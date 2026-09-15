/* =============================================================================
   Analisi dello staging: trascodifica, risoluzione della riga di commessa,
   confronto con l'ultima importazione applicata, calcolo dell'azione.
   NON scrive nulla su L_CMPA / A_PAR.

   Elabora le sole righe con ProcessStatus IN ('PENDING','ERROR'):
   un rilancio non ricalcola cio' che e' gia' DONE o SKIPPED.

   E' QUI che si interviene per cambiare le regole di riconciliazione.
   ============================================================================= */
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

CREATE OR ALTER PROCEDURE dbo.X_sp_ImportPrezzoNetto_Analyze
    @RunId        int,
    @Decimals     tinyint       = 2,      -- arrotondamento del prezzo netto
    @WarnDeltaPct numeric(6,2)  = 30.00   -- soglia di scostamento che genera warning
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;

    IF NOT EXISTS (SELECT 1 FROM dbo.X_ImportPrezzoNetto_Run WHERE RunId = @RunId)
        THROW 50010, 'RunId inesistente.', 1;

    /* -----------------------------------------------------------------------
       1. Normalizzazione dell'importo e rilevazione duplicati nel file.
          COUNT(DISTINCT) non e' disponibile come funzione finestra:
          si confronta MIN con MAX sulla partizione.
       ----------------------------------------------------------------------- */
    IF OBJECT_ID('tempdb..#dup') IS NOT NULL DROP TABLE #dup;

    SELECT r.ExcelRow,
           r.Matricola,
           IsDuplicate = CASE WHEN ROW_NUMBER() OVER (PARTITION BY r.Matricola ORDER BY r.ExcelRow) > 1
                              THEN 1 ELSE 0 END,
           IsAmbiguous = CASE WHEN MIN(ROUND(r.PrezzoNetto, @Decimals)) OVER (PARTITION BY r.Matricola)
                                 <> MAX(ROUND(r.PrezzoNetto, @Decimals)) OVER (PARTITION BY r.Matricola)
                              THEN 1 ELSE 0 END,
           Occurrences = COUNT(*) OVER (PARTITION BY r.Matricola)
    INTO #dup
    FROM dbo.X_ImportPrezzoNetto_Row AS r
    WHERE r.RunId = @RunId
      AND r.Matricola IS NOT NULL
      AND r.PrezzoNetto IS NOT NULL;

    CREATE UNIQUE CLUSTERED INDEX IX_dup ON #dup (ExcelRow);

    /* -----------------------------------------------------------------------
       2. Trascodifica + risoluzione riga di commessa + valori ERP correnti
       ----------------------------------------------------------------------- */
    IF OBJECT_ID('tempdb..#a') IS NOT NULL DROP TABLE #a;

    SELECT r.ExcelRow,
           r.Matricola,
           PrezzoNetto   = ROUND(r.PrezzoNetto, @Decimals),
           r.CommessaExcel,
           r.Modello,
           IsDuplicate   = ISNULL(d.IsDuplicate, 0),
           IsAmbiguous   = ISNULL(d.IsAmbiguous, 0),
           Occurrences   = ISNULL(d.Occurrences, 0),
           t.PACOD,
           t.MatchLevel,
           t.Candidates,
           cm.CONUM,
           cm.IDROW,
           cm.COCOD,
           CommessaRows  = ISNULL(cm.RowCnt, 0),
           OldPAPRZ      = cm.PAPRZ,
           OldPACSA      = p.PACSA,
           PrevRunPrezzo = cur.PrezzoNetto,
           WrittenPAPRZ  = cur.WrittenPAPRZ,
           WrittenPACSA  = cur.WrittenPACSA
    INTO #a
    FROM dbo.X_ImportPrezzoNetto_Row AS r
    LEFT JOIN #dup AS d
           ON d.ExcelRow = r.ExcelRow
    OUTER APPLY dbo.X_fn_trascodeSNpart_Detail(r.Matricola) AS t
    OUTER APPLY (SELECT TOP (1)
                        lc.CONUM,
                        lc.IDROW,
                        COCOD  = RTRIM(ac.COCOD),
                        lc.PAPRZ,
                        RowCnt = COUNT(*) OVER ()
                 FROM dbo.L_CMPA AS lc
                 INNER JOIN dbo.A_COM AS ac ON ac.CONUM = lc.CONUM
                 WHERE lc.PACOD = t.PACOD
                   AND ac.COTYP = 'P'
                 ORDER BY lc.CONUM DESC) AS cm
    LEFT JOIN dbo.A_PAR AS p
           ON p.PACOD = t.PACOD
    LEFT JOIN dbo.X_ImportPrezzoNetto_Current AS cur
           ON cur.Matricola = r.Matricola
    WHERE r.RunId = @RunId
      AND r.ProcessStatus IN ('PENDING', 'ERROR');

    CREATE UNIQUE CLUSTERED INDEX IX_a ON #a (ExcelRow);

    /* -----------------------------------------------------------------------
       3. Calcolo di azione, stato ed eventuali warning
       ----------------------------------------------------------------------- */
    UPDATE r
    SET r.PACOD         = a.PACOD,
        r.MatchLevel    = ISNULL(a.MatchLevel, 'NONE'),
        r.Candidates    = a.Candidates,
        r.CONUM         = a.CONUM,
        r.IDROW         = a.IDROW,
        r.COCOD         = a.COCOD,
        r.OldPAPRZ      = a.OldPAPRZ,
        r.OldPACSA      = a.OldPACSA,
        r.PrevRunPrezzo = a.PrevRunPrezzo,
        r.PrezzoNetto   = a.PrezzoNetto,
        r.Action        = x.Action,
        r.ProcessStatus = CASE
                              WHEN x.Action IN ('NUOVO', 'MODIFICATO', 'RIALLINEATO') THEN 'ANALYZED'
                              WHEN x.Action IN ('AMBIGUO_EXCEL', 'NON_TRASCODIFICATA',
                                                'COMMESSA_NON_TROVATA')               THEN 'ERROR'
                              ELSE 'SKIPPED'
                          END,
        r.ProcessedAt   = CASE
                              WHEN x.Action IN ('NUOVO', 'MODIFICATO', 'RIALLINEATO') THEN NULL
                              ELSE SYSDATETIME()
                          END,
        r.HasWarning    = hw.HasWarning,
        r.Message       = NULLIF(w.Msg, '')
    FROM dbo.X_ImportPrezzoNetto_Row AS r
    INNER JOIN #a AS a
            ON a.ExcelRow = r.ExcelRow
    CROSS APPLY (
        SELECT Action =
            CASE
                WHEN a.Matricola IS NULL                              THEN 'RIGA_IGNORATA'
                WHEN a.IsAmbiguous = 1                                THEN 'AMBIGUO_EXCEL'
                WHEN a.PrezzoNetto IS NULL OR a.PrezzoNetto <= 0      THEN 'SKIPPED_NOPRICE'
                WHEN a.IsDuplicate = 1                                THEN 'DUPLICATO_IGNORATO'
                WHEN a.PACOD IS NULL                                  THEN 'NON_TRASCODIFICATA'
                WHEN a.CONUM IS NULL                                  THEN 'COMMESSA_NON_TROVATA'
                WHEN a.PrevRunPrezzo IS NULL                          THEN 'NUOVO'
                WHEN a.PrezzoNetto <> ROUND(a.PrevRunPrezzo, @Decimals) THEN 'MODIFICATO'
                WHEN ISNULL(a.OldPAPRZ, -1) <> ISNULL(a.WrittenPAPRZ, -1)
                  OR ISNULL(a.OldPACSA, -1) <> ISNULL(a.WrittenPACSA, -1) THEN 'RIALLINEATO'
                ELSE 'INVARIATO'
            END
    ) AS x
    CROSS APPLY (
        /* Normalizzazioni per i controlli incrociati.
           Il modello sull'Excel e il prefisso del codice parte non coincidono
           mai alla lettera: l'ERP aggiunge o omette sigle di allestimento
           (S43/340T2 -> S43/340T2LPC, D24/175T2 <- D24/175T2 FISSO) e usa la
           barra in modo incoerente. Si segnala solo quando nessuno dei due e'
           prefisso dell'altro, cioe' quando si parla di macchine diverse. */
        SELECT ModelNorm   = UPPER(REPLACE(REPLACE(ISNULL(a.Modello, ''), ' ', ''), '/', '')),
               PacodPrefix = UPPER(REPLACE(CASE WHEN a.PACOD IS NULL THEN ''
                                                ELSE LEFT(a.PACOD, LEN(a.PACOD) - CHARINDEX('-', REVERSE(a.PACOD)))
                                           END, '/', '')),
               DeltaPct    = CASE WHEN a.OldPAPRZ > 0 AND a.PrezzoNetto > 0
                                  THEN ROUND(ABS(a.PrezzoNetto - a.OldPAPRZ) * 100.0 / a.OldPAPRZ, 1) END
    ) AS n
    CROSS APPLY (
        SELECT ModelloIncoerente =
            CASE WHEN LEN(n.ModelNorm) > 0 AND LEN(n.PacodPrefix) > 0
                   AND LEFT(n.ModelNorm,   CASE WHEN LEN(n.ModelNorm) < LEN(n.PacodPrefix) THEN LEN(n.ModelNorm) ELSE LEN(n.PacodPrefix) END)
                    <> LEFT(n.PacodPrefix, CASE WHEN LEN(n.ModelNorm) < LEN(n.PacodPrefix) THEN LEN(n.ModelNorm) ELSE LEN(n.PacodPrefix) END)
                 THEN 1 ELSE 0 END
    ) AS mc
    CROSS APPLY (
        SELECT Msg = LTRIM(
                 CASE WHEN a.MatchLevel = 'CORE'
                      THEN 'Trascodifica sul solo nucleo numerico (suffisso lettera diverso fra Excel ed ERP). '
                      ELSE '' END
               + CASE WHEN a.Candidates > 1
                      THEN CONCAT('Parti candidate: ', a.Candidates, ', scelta ', a.PACOD, '. ')
                      ELSE '' END
               + CASE WHEN a.CommessaRows > 1
                      THEN CONCAT('Righe di commessa cliente: ', a.CommessaRows, ', scelta ', a.COCOD, '. ')
                      ELSE '' END
               + CASE WHEN a.CommessaExcel IS NOT NULL AND a.COCOD IS NOT NULL
                        AND TRY_CAST(SUBSTRING(
                                SUBSTRING(a.COCOD, PATINDEX('%[0-9]%', a.COCOD + '0'), 20), 3, 18) AS int)
                            <> a.CommessaExcel
                      THEN CONCAT('Commessa Excel ', a.CommessaExcel, ' non coerente con ', a.COCOD, '. ')
                      ELSE '' END
               + CASE WHEN mc.ModelloIncoerente = 1
                      THEN CONCAT('Modello Excel "', a.Modello, '" non coerente con ', a.PACOD, '. ')
                      ELSE '' END
               + CASE WHEN n.DeltaPct > @WarnDeltaPct
                      THEN CONCAT('Scostamento ',
                                  CAST(CAST(n.DeltaPct AS decimal(12,1)) AS varchar(20)),
                                  '% rispetto al prezzo attuale in ERP. ')
                      ELSE '' END
               + CASE WHEN a.IsAmbiguous = 1
                      THEN CONCAT('La matricola compare ', a.Occurrences,
                                  ' volte nel file con prezzi diversi: nessuna scrittura. ')
                      ELSE '' END
               + CASE WHEN a.IsDuplicate = 1 AND a.IsAmbiguous = 0
                      THEN 'Riga duplicata nel file, stesso prezzo: elaborata la prima occorrenza. '
                      ELSE '' END
               + CASE WHEN a.PACOD IS NULL AND a.Matricola IS NOT NULL AND a.PrezzoNetto > 0
                      THEN 'Nessuna parte in A_PAR con suffisso corrispondente alla matricola. '
                      ELSE '' END
               + CASE WHEN a.PACOD IS NOT NULL AND a.CONUM IS NULL
                      THEN CONCAT('Parte ', a.PACOD, ' senza riga in una commessa cliente (COTYP=''P''). ')
                      ELSE '' END)
    ) AS w
    CROSS APPLY (
        SELECT HasWarning = CASE WHEN LEN(w.Msg) > 0 THEN CAST(1 AS bit) ELSE CAST(0 AS bit) END
    ) AS hw
    WHERE r.RunId = @RunId;

    /* -----------------------------------------------------------------------
       4. Aggiornamento dei contatori di riepilogo
       ----------------------------------------------------------------------- */
    UPDATE run
    SET RowsRead      = s.RowsRead,
        RowsWithPrice = s.RowsWithPrice,
        RowsMatched   = s.RowsMatched,
        RowsUpdated   = s.RowsToUpdate,
        RowsUnchanged = s.RowsUnchanged,
        RowsWarning   = s.RowsWarning,
        RowsError     = s.RowsError
    FROM dbo.X_ImportPrezzoNetto_Run AS run
    CROSS APPLY (
        SELECT RowsRead      = COUNT(*),
               RowsWithPrice = SUM(CASE WHEN PrezzoNetto > 0 THEN 1 ELSE 0 END),
               RowsMatched   = SUM(CASE WHEN PACOD IS NOT NULL THEN 1 ELSE 0 END),
               RowsToUpdate  = SUM(CASE WHEN ProcessStatus IN ('ANALYZED', 'DONE') THEN 1 ELSE 0 END),
               RowsUnchanged = SUM(CASE WHEN Action = 'INVARIATO' THEN 1 ELSE 0 END),
               RowsWarning   = SUM(CASE WHEN HasWarning = 1 THEN 1 ELSE 0 END),
               RowsError     = SUM(CASE WHEN ProcessStatus = 'ERROR' THEN 1 ELSE 0 END)
        FROM dbo.X_ImportPrezzoNetto_Row
        WHERE RunId = @RunId
    ) AS s
    WHERE run.RunId = @RunId;

    /* -----------------------------------------------------------------------
       5. Result set per la griglia di anteprima
       ----------------------------------------------------------------------- */
    EXEC dbo.X_sp_ImportPrezzoNetto_GetRows @RunId = @RunId;
END
GO

PRINT 'X_sp_ImportPrezzoNetto_Analyze : creata.';
GO
