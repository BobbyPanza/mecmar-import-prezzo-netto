/* =============================================================================
   Fotografia di sicurezza PRIMA di un'importazione.

   Copia i valori attuali dei campi che l'import puo' toccare, per tutte le
   macchine (parti con codice <MODELLO>-<MATRICOLA>) e per tutte le righe di
   commessa cliente che le riferiscono - non solo per quelle dell'import in
   corso. Cosi' il ripristino resta possibile anche a distanza di tempo e
   indipendentemente dal log delle singole esecuzioni.

   Ogni esecuzione aggiunge uno snapshot etichettato: nulla viene sovrascritto.

   Uso:
       EXEC dbo.X_sp_ImportPrezzoNetto_Backup @Label = 'prima import 2026-09-15';
   ============================================================================= */
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

IF OBJECT_ID('dbo.X_ImportPrezzoNetto_BackupHead', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.X_ImportPrezzoNetto_BackupHead
    (
        BackupId   int IDENTITY(1,1) NOT NULL CONSTRAINT PK_XIPN_BackupHead PRIMARY KEY,
        TakenAt    datetime2(3)  NOT NULL CONSTRAINT DF_XIPN_BackupHead_TakenAt DEFAULT (SYSDATETIME()),
        TakenBy    nvarchar(128) NULL,
        Label      nvarchar(200) NULL,
        RunId      int           NULL,
        RowsCMPA   int           NOT NULL CONSTRAINT DF_XIPN_BackupHead_RowsCMPA DEFAULT (0),
        RowsPAR    int           NOT NULL CONSTRAINT DF_XIPN_BackupHead_RowsPAR  DEFAULT (0)
    );
END
GO

IF OBJECT_ID('dbo.X_ImportPrezzoNetto_BackupCMPA', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.X_ImportPrezzoNetto_BackupCMPA
    (
        BackupId        int           NOT NULL,
        CONUM           int           NOT NULL,
        IDROW           int           NOT NULL,
        COCOD           varchar(15)   NULL,
        PACOD           varchar(20)   NULL,
        PAPRZ           numeric(18,6) NULL,
        PriceLastUpdate date          NULL,
        CONSTRAINT PK_XIPN_BackupCMPA PRIMARY KEY CLUSTERED (BackupId, CONUM, IDROW)
    );
END
GO

IF OBJECT_ID('dbo.X_ImportPrezzoNetto_BackupPAR', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.X_ImportPrezzoNetto_BackupPAR
    (
        BackupId int           NOT NULL,
        PACOD    varchar(20)   NOT NULL,
        PACSA    numeric(18,6) NULL,
        CONSTRAINT PK_XIPN_BackupPAR PRIMARY KEY CLUSTERED (BackupId, PACOD)
    );
END
GO

CREATE OR ALTER PROCEDURE dbo.X_sp_ImportPrezzoNetto_Backup
    @Label    nvarchar(200) = NULL,
    @RunId    int           = NULL,
    @BackupId int           = NULL OUTPUT
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;

    BEGIN TRANSACTION;

    INSERT INTO dbo.X_ImportPrezzoNetto_BackupHead (TakenBy, Label, RunId)
    VALUES (SUSER_SNAME(), @Label, @RunId);

    SET @BackupId = CAST(SCOPE_IDENTITY() AS int);

    /* Righe di commessa cliente delle macchine */
    INSERT INTO dbo.X_ImportPrezzoNetto_BackupCMPA (BackupId, CONUM, IDROW, COCOD, PACOD, PAPRZ, PriceLastUpdate)
    SELECT @BackupId, lc.CONUM, lc.IDROW, RTRIM(ac.COCOD), RTRIM(lc.PACOD), lc.PAPRZ, lc.PriceLastUpdate
    FROM dbo.L_CMPA AS lc
    INNER JOIN dbo.A_COM AS ac ON ac.CONUM = lc.CONUM
    WHERE ac.COTYP = 'P'
      AND (lc.PACOD LIKE '%-2[0-9][0-9][0-9][0-9][0-9]'
        OR lc.PACOD LIKE '%-2[0-9][0-9][0-9][0-9][0-9][A-Z]');

    /* Anagrafica delle macchine */
    INSERT INTO dbo.X_ImportPrezzoNetto_BackupPAR (BackupId, PACOD, PACSA)
    SELECT @BackupId, RTRIM(p.PACOD), p.PACSA
    FROM dbo.A_PAR AS p
    WHERE p.PACOD LIKE '%-2[0-9][0-9][0-9][0-9][0-9]'
       OR p.PACOD LIKE '%-2[0-9][0-9][0-9][0-9][0-9][A-Z]';

    UPDATE h
    SET RowsCMPA = (SELECT COUNT(*) FROM dbo.X_ImportPrezzoNetto_BackupCMPA WHERE BackupId = @BackupId),
        RowsPAR  = (SELECT COUNT(*) FROM dbo.X_ImportPrezzoNetto_BackupPAR  WHERE BackupId = @BackupId)
    FROM dbo.X_ImportPrezzoNetto_BackupHead AS h
    WHERE h.BackupId = @BackupId;

    COMMIT TRANSACTION;

    SELECT BackupId, TakenAt, Label, RowsCMPA, RowsPAR
    FROM dbo.X_ImportPrezzoNetto_BackupHead
    WHERE BackupId = @BackupId;
END
GO

/* -----------------------------------------------------------------------------
   Ripristino integrale da uno snapshot.
   Rete di sicurezza di ultima istanza: riporta TUTTE le righe fotografate ai
   valori del backup, non solo quelle di una singola importazione.
   ----------------------------------------------------------------------------- */
CREATE OR ALTER PROCEDURE dbo.X_sp_ImportPrezzoNetto_RestoreBackup
    @BackupId int,
    @WhatIf   bit = 1
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;

    IF NOT EXISTS (SELECT 1 FROM dbo.X_ImportPrezzoNetto_BackupHead WHERE BackupId = @BackupId)
        THROW 50040, 'BackupId inesistente.', 1;

    IF @WhatIf = 1
    BEGIN
        SELECT Tabella = 'L_CMPA', b.CONUM, b.IDROW, b.PACOD,
               ValoreAttuale = lc.PAPRZ, ValoreDaRipristinare = b.PAPRZ
        FROM dbo.X_ImportPrezzoNetto_BackupCMPA AS b
        INNER JOIN dbo.L_CMPA AS lc ON lc.CONUM = b.CONUM AND lc.IDROW = b.IDROW
        WHERE b.BackupId = @BackupId
          AND ISNULL(lc.PAPRZ, -1) <> ISNULL(b.PAPRZ, -1);

        SELECT Tabella = 'A_PAR', b.PACOD,
               ValoreAttuale = p.PACSA, ValoreDaRipristinare = b.PACSA
        FROM dbo.X_ImportPrezzoNetto_BackupPAR AS b
        INNER JOIN dbo.A_PAR AS p ON p.PACOD = b.PACOD
        WHERE b.BackupId = @BackupId
          AND ISNULL(p.PACSA, -1) <> ISNULL(b.PACSA, -1);

        PRINT 'Simulazione: rilanciare con @WhatIf = 0 per ripristinare davvero.';
        RETURN;
    END

    BEGIN TRANSACTION;

    UPDATE lc
    SET lc.PAPRZ           = b.PAPRZ,
        lc.PriceLastUpdate = b.PriceLastUpdate
    FROM dbo.L_CMPA AS lc
    INNER JOIN dbo.X_ImportPrezzoNetto_BackupCMPA AS b
            ON b.CONUM = lc.CONUM AND b.IDROW = lc.IDROW
    WHERE b.BackupId = @BackupId;

    UPDATE p
    SET p.PACSA = b.PACSA
    FROM dbo.A_PAR AS p
    INNER JOIN dbo.X_ImportPrezzoNetto_BackupPAR AS b
            ON b.PACOD = p.PACOD
    WHERE b.BackupId = @BackupId;

    COMMIT TRANSACTION;

    SELECT Esito = 'Ripristino completato.', BackupId = @BackupId;
END
GO

PRINT 'X_ImportPrezzoNetto_Backup* : tabelle e procedure create.';
GO
