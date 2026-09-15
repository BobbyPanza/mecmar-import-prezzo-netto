/* =============================================================================
   Annullamento di un'importazione: rete di sicurezza.

   Ripristina su L_CMPA e A_PAR i valori registrati in X_ImportPrezzoNetto_Log
   per il run indicato, riporta X_ImportPrezzoNetto_Current allo stato
   precedente e marca il run come ROLLEDBACK.

   Attenzione: se dopo l'import qualcuno ha modificato a mano i prezzi,
   il rollback sovrascrive anche quelle modifiche. Usare @Force = 1 per
   procedere comunque; senza, la procedura si ferma e li elenca.
   ============================================================================= */
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

CREATE OR ALTER PROCEDURE dbo.X_sp_ImportPrezzoNetto_Rollback
    @RunId int,
    @Force bit           = 0,
    @By    nvarchar(128) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;

    SET @By = ISNULL(@By, SUSER_SNAME());

    IF NOT EXISTS (SELECT 1 FROM dbo.X_ImportPrezzoNetto_Run WHERE RunId = @RunId)
        THROW 50030, 'RunId inesistente.', 1;

    IF NOT EXISTS (SELECT 1 FROM dbo.X_ImportPrezzoNetto_Log
                   WHERE RunId = @RunId AND RevertedByRunId IS NULL)
        THROW 50031, 'Nessuna scrittura da annullare per questo run.', 1;

    /* --- controllo: valori cambiati a mano dopo l'import ----------------- */
    IF OBJECT_ID('tempdb..#drift') IS NOT NULL DROP TABLE #drift;

    SELECT g.LogId, g.Matricola, g.TargetTable, g.TargetColumn, g.TargetKey,
           ValoreScrittoDaNoi = g.NewValue,
           ValoreAttuale      = CASE g.TargetTable WHEN 'L_CMPA' THEN lc.PAPRZ ELSE p.PACSA END
    INTO #drift
    FROM dbo.X_ImportPrezzoNetto_Log AS g
    LEFT JOIN dbo.L_CMPA AS lc
           ON g.TargetTable = 'L_CMPA'
          AND lc.CONUM = TRY_CAST(SUBSTRING(g.TargetKey, 7, CHARINDEX(';', g.TargetKey) - 7) AS int)
          AND lc.IDROW = TRY_CAST(SUBSTRING(g.TargetKey, CHARINDEX('IDROW=', g.TargetKey) + 6, 20) AS int)
    LEFT JOIN dbo.A_PAR AS p
           ON g.TargetTable = 'A_PAR'
          AND p.PACOD = SUBSTRING(g.TargetKey, 7, 20)
    WHERE g.RunId = @RunId
      AND g.RevertedByRunId IS NULL
      AND g.TargetColumn IN ('PAPRZ', 'PACSA')
      AND ISNULL(CASE g.TargetTable WHEN 'L_CMPA' THEN lc.PAPRZ ELSE p.PACSA END, -1)
          <> ISNULL(g.NewValue, -1);

    IF EXISTS (SELECT 1 FROM #drift) AND @Force = 0
    BEGIN
        SELECT Avviso = 'Valori modificati dopo l''importazione: rilanciare con @Force = 1 per sovrascriverli.', *
        FROM #drift;
        RETURN;
    END

    BEGIN TRY
        BEGIN TRANSACTION;

        /* --- ripristino L_CMPA.PAPRZ ------------------------------------- */
        UPDATE lc
        SET lc.PAPRZ = g.OldValue
        FROM dbo.L_CMPA AS lc
        INNER JOIN dbo.X_ImportPrezzoNetto_Log AS g
                ON g.TargetTable = 'L_CMPA' AND g.TargetColumn = 'PAPRZ'
               AND lc.CONUM = TRY_CAST(SUBSTRING(g.TargetKey, 7, CHARINDEX(';', g.TargetKey) - 7) AS int)
               AND lc.IDROW = TRY_CAST(SUBSTRING(g.TargetKey, CHARINDEX('IDROW=', g.TargetKey) + 6, 20) AS int)
        WHERE g.RunId = @RunId AND g.RevertedByRunId IS NULL;

        /* --- ripristino L_CMPA.PriceLastUpdate --------------------------- */
        UPDATE lc
        SET lc.PriceLastUpdate = g.OldValueDate
        FROM dbo.L_CMPA AS lc
        INNER JOIN dbo.X_ImportPrezzoNetto_Log AS g
                ON g.TargetTable = 'L_CMPA' AND g.TargetColumn = 'PriceLastUpdate'
               AND lc.CONUM = TRY_CAST(SUBSTRING(g.TargetKey, 7, CHARINDEX(';', g.TargetKey) - 7) AS int)
               AND lc.IDROW = TRY_CAST(SUBSTRING(g.TargetKey, CHARINDEX('IDROW=', g.TargetKey) + 6, 20) AS int)
        WHERE g.RunId = @RunId AND g.RevertedByRunId IS NULL;

        /* --- ripristino A_PAR.PACSA -------------------------------------- */
        UPDATE p
        SET p.PACSA = g.OldValue
        FROM dbo.A_PAR AS p
        INNER JOIN dbo.X_ImportPrezzoNetto_Log AS g
                ON g.TargetTable = 'A_PAR' AND g.TargetColumn = 'PACSA'
               AND p.PACOD = SUBSTRING(g.TargetKey, 7, 20)
        WHERE g.RunId = @RunId AND g.RevertedByRunId IS NULL;

        /* --- ripristino dello stato "ultima importazione" ---------------- */
        DELETE cur
        FROM dbo.X_ImportPrezzoNetto_Current AS cur
        INNER JOIN dbo.X_ImportPrezzoNetto_Row AS r
                ON r.Matricola = cur.Matricola
        WHERE r.RunId = @RunId AND r.Applied = 1 AND r.PrevRunPrezzo IS NULL
          AND cur.SourceRunId = @RunId;

        UPDATE cur
        SET cur.PrezzoNetto  = r.PrevRunPrezzo,
            cur.WrittenPAPRZ = r.OldPAPRZ,
            cur.WrittenPACSA = r.OldPACSA,
            cur.UpdatedAt    = SYSDATETIME(),
            cur.UpdatedBy    = @By
        FROM dbo.X_ImportPrezzoNetto_Current AS cur
        INNER JOIN dbo.X_ImportPrezzoNetto_Row AS r
                ON r.Matricola = cur.Matricola
        WHERE r.RunId = @RunId AND r.Applied = 1 AND r.PrevRunPrezzo IS NOT NULL
          AND cur.SourceRunId = @RunId;

        /* --- marcatura ---------------------------------------------------- */
        UPDATE dbo.X_ImportPrezzoNetto_Log
        SET RevertedByRunId = @RunId
        WHERE RunId = @RunId AND RevertedByRunId IS NULL;

        UPDATE dbo.X_ImportPrezzoNetto_Row
        SET ProcessStatus = 'ANALYZED',
            Applied       = 0,
            Message       = CONCAT(N'Importazione annullata il ',
                                   CONVERT(varchar(19), SYSDATETIME(), 120), N'. ', Message)
        WHERE RunId = @RunId AND Applied = 1;

        UPDATE dbo.X_ImportPrezzoNetto_Run
        SET Status = 'ROLLEDBACK'
        WHERE RunId = @RunId;

        COMMIT TRANSACTION;
    END TRY
    BEGIN CATCH
        IF XACT_STATE() <> 0 ROLLBACK TRANSACTION;
        THROW;
    END CATCH

    SELECT Esito = 'Rollback completato.', RunId = @RunId;
END
GO

PRINT 'X_sp_ImportPrezzoNetto_Rollback : creata.';
GO
