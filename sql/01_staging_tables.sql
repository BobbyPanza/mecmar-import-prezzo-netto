/* =============================================================================
   Import Prezzo Netto - Tabelle di staging e storico
   DB: Factory   Schema: dbo   Prefisso: X_ImportPrezzoNetto_
   Script idempotente: puo' essere rieseguito senza perdere dati.
   ============================================================================= */
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

/* ---------------------------------------------------------------------------
   X_ImportPrezzoNetto_Run : una riga per esecuzione
   --------------------------------------------------------------------------- */
IF OBJECT_ID('dbo.X_ImportPrezzoNetto_Run', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.X_ImportPrezzoNetto_Run
    (
        RunId                int            IDENTITY(1,1) NOT NULL,
        StartedAt            datetime2(3)   NOT NULL CONSTRAINT DF_XIPN_Run_StartedAt DEFAULT (SYSDATETIME()),
        FinishedAt           datetime2(3)   NULL,
        SourceFile           nvarchar(500)  NOT NULL,
        SourceFileHash       char(64)       NULL,
        SourceFileModifiedAt datetime2(3)   NULL,
        SheetName            nvarchar(128)  NOT NULL,
        Mode                 varchar(10)    NOT NULL,   -- APPLY | DRYRUN
        Status               varchar(20)    NOT NULL,   -- RUNNING | COMPLETED | FAILED | ROLLEDBACK
        UserName             nvarchar(128)  NULL,
        MachineName          nvarchar(128)  NULL,
        RowsRead             int            NOT NULL CONSTRAINT DF_XIPN_Run_RowsRead      DEFAULT (0),
        RowsWithPrice        int            NOT NULL CONSTRAINT DF_XIPN_Run_RowsWithPrice DEFAULT (0),
        RowsMatched          int            NOT NULL CONSTRAINT DF_XIPN_Run_RowsMatched   DEFAULT (0),
        RowsUpdated          int            NOT NULL CONSTRAINT DF_XIPN_Run_RowsUpdated   DEFAULT (0),
        RowsUnchanged        int            NOT NULL CONSTRAINT DF_XIPN_Run_RowsUnchanged DEFAULT (0),
        RowsWarning          int            NOT NULL CONSTRAINT DF_XIPN_Run_RowsWarning   DEFAULT (0),
        RowsError            int            NOT NULL CONSTRAINT DF_XIPN_Run_RowsError     DEFAULT (0),
        ErrorMessage         nvarchar(max)  NULL,
        CONSTRAINT PK_X_ImportPrezzoNetto_Run PRIMARY KEY CLUSTERED (RunId)
    );

    CREATE INDEX IX_XIPN_Run_Hash   ON dbo.X_ImportPrezzoNetto_Run (SourceFileHash, Status);
    CREATE INDEX IX_XIPN_Run_Status ON dbo.X_ImportPrezzoNetto_Run (Status, StartedAt DESC);
END
GO

/* ---------------------------------------------------------------------------
   X_ImportPrezzoNetto_Row : staging, una riga per riga dell'Excel
   --------------------------------------------------------------------------- */
IF OBJECT_ID('dbo.X_ImportPrezzoNetto_Row', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.X_ImportPrezzoNetto_Row
    (
        RunId           int            NOT NULL,
        ExcelRow        int            NOT NULL,

        -- dati letti dal foglio
        MatricolaRaw    nvarchar(50)   NULL,
        Matricola       varchar(30)    NULL,
        PrezzoNetto     numeric(18,6)  NULL,
        CommessaExcel   int            NULL,
        Modello         nvarchar(160)  NULL,
        Cliente         nvarchar(160)  NULL,
        DataConsegna    date           NULL,

        -- esito della trascodifica
        PACOD           varchar(20)    NULL,
        MatchLevel      varchar(10)    NULL,   -- EXACT | CORE | NONE
        Candidates      int            NULL,

        -- riga di commessa risolta
        CONUM           int            NULL,
        IDROW           int            NULL,
        COCOD           varchar(15)    NULL,

        -- valori ERP prima della scrittura
        OldPAPRZ        numeric(18,6)  NULL,
        OldPACSA        numeric(18,6)  NULL,
        PrevRunPrezzo   numeric(18,6)  NULL,

        -- esito
        Action          varchar(30)    NULL,
        HasWarning      bit            NOT NULL CONSTRAINT DF_XIPN_Row_HasWarning DEFAULT (0),
        ProcessStatus   varchar(15)    NOT NULL CONSTRAINT DF_XIPN_Row_ProcessStatus DEFAULT ('PENDING'),
        ProcessedAt     datetime2(3)   NULL,
        ProcessAttempts int            NOT NULL CONSTRAINT DF_XIPN_Row_Attempts DEFAULT (0),
        Applied         bit            NOT NULL CONSTRAINT DF_XIPN_Row_Applied DEFAULT (0),
        Message         nvarchar(500)  NULL,

        CONSTRAINT PK_X_ImportPrezzoNetto_Row PRIMARY KEY CLUSTERED (RunId, ExcelRow),
        CONSTRAINT FK_XIPN_Row_Run FOREIGN KEY (RunId)
            REFERENCES dbo.X_ImportPrezzoNetto_Run (RunId),
        CONSTRAINT CK_XIPN_Row_ProcessStatus CHECK
            (ProcessStatus IN ('PENDING','ANALYZED','DONE','SKIPPED','ERROR'))
    );

    CREATE INDEX IX_XIPN_Row_Pending ON dbo.X_ImportPrezzoNetto_Row (RunId, ProcessStatus)
        INCLUDE (Action, Matricola);
    CREATE INDEX IX_XIPN_Row_Matricola ON dbo.X_ImportPrezzoNetto_Row (Matricola, RunId);
END
GO

/* ---------------------------------------------------------------------------
   X_ImportPrezzoNetto_Current : stato dell'ultima importazione applicata
   --------------------------------------------------------------------------- */
IF OBJECT_ID('dbo.X_ImportPrezzoNetto_Current', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.X_ImportPrezzoNetto_Current
    (
        Matricola     varchar(30)    NOT NULL,
        PACOD         varchar(20)    NULL,
        CONUM         int            NULL,
        IDROW         int            NULL,
        COCOD         varchar(15)    NULL,
        PrezzoNetto   numeric(18,6)  NOT NULL,
        WrittenPAPRZ  numeric(18,6)  NULL,
        WrittenPACSA  numeric(18,6)  NULL,
        SourceRunId   int            NOT NULL,
        UpdatedAt     datetime2(3)   NOT NULL CONSTRAINT DF_XIPN_Cur_UpdatedAt DEFAULT (SYSDATETIME()),
        UpdatedBy     nvarchar(128)  NULL,
        CONSTRAINT PK_X_ImportPrezzoNetto_Current PRIMARY KEY CLUSTERED (Matricola)
    );
END
GO

/* ---------------------------------------------------------------------------
   X_ImportPrezzoNetto_Log : audit di ogni singola scrittura su ERP
   --------------------------------------------------------------------------- */
IF OBJECT_ID('dbo.X_ImportPrezzoNetto_Log', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.X_ImportPrezzoNetto_Log
    (
        LogId               bigint         IDENTITY(1,1) NOT NULL,
        RunId               int            NOT NULL,
        ExcelRow            int            NULL,
        Matricola           varchar(30)    NULL,
        TargetTable         varchar(30)    NOT NULL,   -- L_CMPA | A_PAR
        TargetColumn        varchar(30)    NOT NULL,   -- PAPRZ | PriceLastUpdate | PACSA
        TargetKey           nvarchar(100)  NOT NULL,
        OldValue            numeric(18,6)  NULL,
        NewValue            numeric(18,6)  NULL,
        OldValueDate        date           NULL,
        NewValueDate        date           NULL,
        ManualChangeDetected bit           NOT NULL CONSTRAINT DF_XIPN_Log_Manual DEFAULT (0),
        RevertedByRunId     int            NULL,
        WrittenAt           datetime2(3)   NOT NULL CONSTRAINT DF_XIPN_Log_WrittenAt DEFAULT (SYSDATETIME()),
        WrittenBy           nvarchar(128)  NULL,
        CONSTRAINT PK_X_ImportPrezzoNetto_Log PRIMARY KEY CLUSTERED (LogId),
        CONSTRAINT FK_XIPN_Log_Run FOREIGN KEY (RunId)
            REFERENCES dbo.X_ImportPrezzoNetto_Run (RunId)
    );

    CREATE INDEX IX_XIPN_Log_Run ON dbo.X_ImportPrezzoNetto_Log (RunId, LogId);
END
GO

PRINT 'X_ImportPrezzoNetto_* : tabelle verificate/create.';
GO
