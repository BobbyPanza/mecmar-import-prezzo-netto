/* =============================================================================
   Applicazione dello staging all'ERP.

   >>> E' QUESTA LA PROCEDURA DA MODIFICARE PER CAMBIARE COSA SI SCRIVE. <<<

   Campi aggiornati oggi:
       L_CMPA.PAPRZ            prezzo unitario della riga di commessa cliente
       L_CMPA.PriceLastUpdate  data di applicazione dell'import
       A_PAR.PACSA             prezzo sulla parte <MODELLO>-<MATRICOLA>

   Tutto avviene in una sola transazione: o passa tutto o non passa niente.
   Ogni scrittura lascia una riga in X_ImportPrezzoNetto_Log con il valore
   precedente, quindi e' sempre annullabile con X_sp_ImportPrezzoNetto_Rollback.

   Rilanciabile: elabora le sole righe ANALYZED/ERROR, ignora le DONE.
   ============================================================================= */
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

CREATE OR ALTER PROCEDURE dbo.X_sp_ImportPrezzoNetto_Apply
    @RunId     int,
    @DryRun    bit           = 0,
    @AppliedBy nvarchar(128) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;

    DECLARE @Today date = CAST(SYSDATETIME() AS date);
    SET @AppliedBy = ISNULL(@AppliedBy, SUSER_SNAME());

    DECLARE @Status varchar(20);
    SELECT @Status = Status FROM dbo.X_ImportPrezzoNetto_Run WHERE RunId = @RunId;

    IF @Status IS NULL
        THROW 50020, 'RunId inesistente.', 1;
    IF @Status <> 'RUNNING'
        THROW 50021, 'Il run non e'' piu'' in stato RUNNING: non e'' applicabile.', 1;

    /* -----------------------------------------------------------------------
       Insieme delle righe da scrivere, congelato prima di toccare l'ERP.
       ----------------------------------------------------------------------- */
    IF OBJECT_ID('tempdb..#apply') IS NOT NULL DROP TABLE #apply;

    SELECT r.ExcelRow,
           r.Matricola,
           r.PACOD,
           r.CONUM,
           r.IDROW,
           r.COCOD,
           r.PrezzoNetto,
           r.PrevRunPrezzo,
           OldPAPRZ            = lc.PAPRZ,
           OldPriceLastUpdate  = lc.PriceLastUpdate,
           OldPACSA            = p.PACSA,
           ManualChange        = CASE
                                     WHEN cur.Matricola IS NULL THEN 0
                                     WHEN ISNULL(lc.PAPRZ, -1) <> ISNULL(cur.WrittenPAPRZ, -1)
                                       OR ISNULL(p.PACSA, -1)  <> ISNULL(cur.WrittenPACSA, -1) THEN 1
                                     ELSE 0
                                 END
    INTO #apply
    FROM dbo.X_ImportPrezzoNetto_Row AS r
    INNER JOIN dbo.L_CMPA AS lc
            ON lc.CONUM = r.CONUM AND lc.IDROW = r.IDROW
    INNER JOIN dbo.A_PAR AS p
            ON p.PACOD = r.PACOD
    LEFT JOIN dbo.X_ImportPrezzoNetto_Current AS cur
            ON cur.Matricola = r.Matricola
    WHERE r.RunId = @RunId
      AND r.ProcessStatus IN ('ANALYZED', 'ERROR')
      AND r.Action IN ('NUOVO', 'MODIFICATO', 'RIALLINEATO')
      AND r.PrezzoNetto > 0;

    CREATE UNIQUE CLUSTERED INDEX IX_apply ON #apply (ExcelRow);

    /* -----------------------------------------------------------------------
       Simulazione: nessuna scrittura, si restituisce solo cosa si farebbe.
       ----------------------------------------------------------------------- */
    IF @DryRun = 1
    BEGIN
        SELECT Simulazione = CAST(1 AS bit), * FROM #apply ORDER BY ExcelRow;
        RETURN;
    END

    /* -----------------------------------------------------------------------
       Righe di staging che puntano a una riga di commessa sparita nel frattempo
       ----------------------------------------------------------------------- */
    UPDATE r
    SET r.ProcessStatus   = 'ERROR',
        r.ProcessAttempts = r.ProcessAttempts + 1,
        r.ProcessedAt     = SYSDATETIME(),
        r.Message         = CONCAT(N'Riga di commessa non piu'' presente in ERP (CONUM=',
                                   r.CONUM, N', IDROW=', r.IDROW, N'). ', r.Message)
    FROM dbo.X_ImportPrezzoNetto_Row AS r
    WHERE r.RunId = @RunId
      AND r.ProcessStatus IN ('ANALYZED', 'ERROR')
      AND r.Action IN ('NUOVO', 'MODIFICATO', 'RIALLINEATO')
      AND NOT EXISTS (SELECT 1 FROM #apply AS a WHERE a.ExcelRow = r.ExcelRow);

    BEGIN TRY
        BEGIN TRANSACTION;

        /* --- audit: L_CMPA.PAPRZ ---------------------------------------- */
        INSERT INTO dbo.X_ImportPrezzoNetto_Log
            (RunId, ExcelRow, Matricola, TargetTable, TargetColumn, TargetKey,
             OldValue, NewValue, ManualChangeDetected, WrittenBy)
        SELECT @RunId, a.ExcelRow, a.Matricola, 'L_CMPA', 'PAPRZ',
               CONCAT('CONUM=', a.CONUM, ';IDROW=', a.IDROW),
               a.OldPAPRZ, a.PrezzoNetto, a.ManualChange, @AppliedBy
        FROM #apply AS a;

        /* --- audit: L_CMPA.PriceLastUpdate ------------------------------ */
        INSERT INTO dbo.X_ImportPrezzoNetto_Log
            (RunId, ExcelRow, Matricola, TargetTable, TargetColumn, TargetKey,
             OldValueDate, NewValueDate, WrittenBy)
        SELECT @RunId, a.ExcelRow, a.Matricola, 'L_CMPA', 'PriceLastUpdate',
               CONCAT('CONUM=', a.CONUM, ';IDROW=', a.IDROW),
               a.OldPriceLastUpdate, @Today, @AppliedBy
        FROM #apply AS a;

        /* --- audit: A_PAR.PACSA ----------------------------------------- */
        INSERT INTO dbo.X_ImportPrezzoNetto_Log
            (RunId, ExcelRow, Matricola, TargetTable, TargetColumn, TargetKey,
             OldValue, NewValue, ManualChangeDetected, WrittenBy)
        SELECT @RunId, a.ExcelRow, a.Matricola, 'A_PAR', 'PACSA',
               CONCAT('PACOD=', a.PACOD),
               a.OldPACSA, a.PrezzoNetto, a.ManualChange, @AppliedBy
        FROM #apply AS a;

        /* --- scrittura: riga di commessa -------------------------------- */
        UPDATE lc
        SET lc.PAPRZ           = a.PrezzoNetto,
            lc.PriceLastUpdate = @Today
        FROM dbo.L_CMPA AS lc
        INNER JOIN #apply AS a
                ON a.CONUM = lc.CONUM AND a.IDROW = lc.IDROW;

        /* --- scrittura: parte ------------------------------------------- */
        UPDATE p
        SET p.PACSA = a.PrezzoNetto
        FROM dbo.A_PAR AS p
        INNER JOIN #apply AS a
                ON a.PACOD = p.PACOD;

        /* --- stato "ultima importazione applicata" ----------------------- */
        MERGE dbo.X_ImportPrezzoNetto_Current AS tgt
        USING #apply AS src
           ON tgt.Matricola = src.Matricola
        WHEN MATCHED THEN UPDATE SET
            tgt.PACOD        = src.PACOD,
            tgt.CONUM        = src.CONUM,
            tgt.IDROW        = src.IDROW,
            tgt.COCOD        = src.COCOD,
            tgt.PrezzoNetto  = src.PrezzoNetto,
            tgt.WrittenPAPRZ = src.PrezzoNetto,
            tgt.WrittenPACSA = src.PrezzoNetto,
            tgt.SourceRunId  = @RunId,
            tgt.UpdatedAt    = SYSDATETIME(),
            tgt.UpdatedBy    = @AppliedBy
        WHEN NOT MATCHED BY TARGET THEN INSERT
            (Matricola, PACOD, CONUM, IDROW, COCOD, PrezzoNetto,
             WrittenPAPRZ, WrittenPACSA, SourceRunId, UpdatedBy)
            VALUES
            (src.Matricola, src.PACOD, src.CONUM, src.IDROW, src.COCOD, src.PrezzoNetto,
             src.PrezzoNetto, src.PrezzoNetto, @RunId, @AppliedBy);

        /* --- chiusura delle righe di staging ----------------------------- */
        UPDATE r
        SET r.ProcessStatus   = 'DONE',
            r.Applied         = 1,
            r.ProcessedAt     = SYSDATETIME(),
            r.ProcessAttempts = r.ProcessAttempts + 1,
            r.OldPAPRZ        = a.OldPAPRZ,
            r.OldPACSA        = a.OldPACSA,
            r.Message         = CASE WHEN a.ManualChange = 1
                                     THEN CONCAT(N'Valore modificato a mano in ERP dopo l''ultima importazione: sovrascritto. ', r.Message)
                                     ELSE r.Message END,
            r.HasWarning      = CASE WHEN a.ManualChange = 1 THEN 1 ELSE r.HasWarning END
        FROM dbo.X_ImportPrezzoNetto_Row AS r
        INNER JOIN #apply AS a ON a.ExcelRow = r.ExcelRow
        WHERE r.RunId = @RunId;

        /* --- chiusura del run -------------------------------------------- */
        UPDATE run
        SET Status      = 'COMPLETED',
            FinishedAt  = SYSDATETIME(),
            RowsUpdated = s.RowsDone,
            RowsError   = s.RowsError,
            RowsWarning = s.RowsWarning
        FROM dbo.X_ImportPrezzoNetto_Run AS run
        CROSS APPLY (
            SELECT RowsDone    = SUM(CASE WHEN ProcessStatus = 'DONE'  THEN 1 ELSE 0 END),
                   RowsError   = SUM(CASE WHEN ProcessStatus = 'ERROR' THEN 1 ELSE 0 END),
                   RowsWarning = SUM(CASE WHEN HasWarning = 1 THEN 1 ELSE 0 END)
            FROM dbo.X_ImportPrezzoNetto_Row WHERE RunId = @RunId
        ) AS s
        WHERE run.RunId = @RunId;

        COMMIT TRANSACTION;
    END TRY
    BEGIN CATCH
        IF XACT_STATE() <> 0 ROLLBACK TRANSACTION;

        UPDATE dbo.X_ImportPrezzoNetto_Run
        SET Status       = 'FAILED',
            FinishedAt   = SYSDATETIME(),
            ErrorMessage = ERROR_MESSAGE()
        WHERE RunId = @RunId;

        THROW;
    END CATCH

    EXEC dbo.X_sp_ImportPrezzoNetto_GetRows @RunId = @RunId;
END
GO

PRINT 'X_sp_ImportPrezzoNetto_Apply : creata.';
GO
