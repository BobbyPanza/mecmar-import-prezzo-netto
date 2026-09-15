/* =============================================================================
   Apre una nuova esecuzione di import e restituisce il RunId.
   Se esiste gia' un run RUNNING sullo stesso file (stesso hash) lo restituisce
   invece di crearne uno nuovo, cosi' l'applicativo puo' proporre la ripartenza.
   ============================================================================= */
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

CREATE OR ALTER PROCEDURE dbo.X_sp_ImportPrezzoNetto_BeginRun
    @SourceFile           nvarchar(500),
    @SheetName            nvarchar(128),
    @Mode                 varchar(10)   = 'APPLY',
    @SourceFileHash       char(64)      = NULL,
    @SourceFileModifiedAt datetime2(3)  = NULL,
    @UserName             nvarchar(128) = NULL,
    @MachineName          nvarchar(128) = NULL,
    @Resume               bit           = 0,
    @ResumeAfterMinutes   int           = 10,
    @RunId                int           = NULL OUTPUT
AS
BEGIN
    SET NOCOUNT ON;

    IF @Mode NOT IN ('APPLY', 'DRYRUN')
        THROW 50001, 'Mode ammessi: APPLY, DRYRUN.', 1;

    SET @RunId = NULL;

    /* Ripartenza da un run rimasto in sospeso sullo stesso file.
       La soglia temporale evita di agganciare un'importazione che un altro
       utente sta eseguendo proprio adesso: un run ancora "fresco" si presume
       vivo, non interrotto. */
    IF @Resume = 1 AND @SourceFileHash IS NOT NULL
    BEGIN
        SELECT TOP (1) @RunId = RunId
        FROM dbo.X_ImportPrezzoNetto_Run
        WHERE SourceFileHash = @SourceFileHash
          AND Status = 'RUNNING'
          AND StartedAt < DATEADD(MINUTE, -@ResumeAfterMinutes, SYSDATETIME())
        ORDER BY RunId DESC;
    END

    IF @RunId IS NULL
    BEGIN
        INSERT INTO dbo.X_ImportPrezzoNetto_Run
            (SourceFile, SourceFileHash, SourceFileModifiedAt, SheetName,
             Mode, Status, UserName, MachineName)
        VALUES
            (@SourceFile, @SourceFileHash, @SourceFileModifiedAt, @SheetName,
             @Mode, 'RUNNING', @UserName, @MachineName);

        SET @RunId = CAST(SCOPE_IDENTITY() AS int);
    END
    ELSE
    BEGIN
        UPDATE dbo.X_ImportPrezzoNetto_Run
        SET Mode        = @Mode,
            SourceFile  = @SourceFile,
            SheetName   = @SheetName,
            UserName    = @UserName,
            MachineName = @MachineName
        WHERE RunId = @RunId;
    END

    SELECT RunId = @RunId,
           Resumed = CASE WHEN @Resume = 1 AND EXISTS (SELECT 1 FROM dbo.X_ImportPrezzoNetto_Row WHERE RunId = @RunId)
                          THEN CAST(1 AS bit) ELSE CAST(0 AS bit) END;
END
GO

PRINT 'X_sp_ImportPrezzoNetto_BeginRun : creata.';
GO
