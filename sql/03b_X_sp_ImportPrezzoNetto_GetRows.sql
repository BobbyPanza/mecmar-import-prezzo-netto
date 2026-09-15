/* =============================================================================
   Result set unico usato dalla griglia WPF e dal report.
   Tenuto separato dalle procedure di analisi/applicazione cosi' che aggiungere
   una colonna alla vista dell'utente non richieda di toccare la logica.
   ============================================================================= */
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

CREATE OR ALTER PROCEDURE dbo.X_sp_ImportPrezzoNetto_GetRows
    @RunId int,
    @OnlyChanges bit = 0,
    @OnlyProblems bit = 0
AS
BEGIN
    SET NOCOUNT ON;

    SELECT  r.ExcelRow,
            r.MatricolaRaw,
            r.Matricola,
            r.Modello,
            r.Cliente,
            r.DataConsegna,
            r.CommessaExcel,
            r.PACOD,
            r.MatchLevel,
            r.Candidates,
            r.COCOD,
            r.CONUM,
            r.IDROW,
            PrezzoErp    = r.OldPAPRZ,
            PrezzoNetto  = r.PrezzoNetto,
            Delta        = CASE WHEN r.OldPAPRZ IS NOT NULL AND r.PrezzoNetto IS NOT NULL
                                THEN r.PrezzoNetto - r.OldPAPRZ END,
            DeltaPct     = CASE WHEN r.OldPAPRZ > 0 AND r.PrezzoNetto IS NOT NULL
                                THEN ROUND((r.PrezzoNetto - r.OldPAPRZ) * 100.0 / r.OldPAPRZ, 2) END,
            PrezzoPrecedenteImport = r.PrevRunPrezzo,
            r.Action,
            r.ProcessStatus,
            r.ProcessAttempts,
            r.ProcessedAt,
            r.Applied,
            r.HasWarning,
            r.Message
    FROM dbo.X_ImportPrezzoNetto_Row AS r
    WHERE r.RunId = @RunId
      AND (@OnlyChanges  = 0 OR r.Action IN ('NUOVO', 'MODIFICATO', 'RIALLINEATO'))
      AND (@OnlyProblems = 0 OR r.ProcessStatus = 'ERROR' OR r.HasWarning = 1)
    ORDER BY r.ExcelRow;
END
GO

PRINT 'X_sp_ImportPrezzoNetto_GetRows : creata.';
GO
