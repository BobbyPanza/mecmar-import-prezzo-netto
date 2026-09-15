/* =============================================================================
   Trascodifica MATRICOLA (serial number) -> codice parte A_PAR.PACOD

   In Factory la macchina finita e' codificata <MODELLO>-<MATRICOLA>
       D20/153THP-226024    MECMAR GRAIN DRYER MOD. D20/153T HP SN 226024
       D24/175T2-226019K    MECMAR GRAIN  DRYER MOD. D24/175T2 SN 226019K
   mentre i componenti di commessa usano il prefisso numerico
       226024TESTATA / 226024CIL-COMP / 226024QE
   e quindi NON intercettano il pattern "suffisso dopo l'ultimo trattino".

   Livelli di match:
     EXACT  il testo dopo l'ultimo '-' coincide con la matricola normalizzata
     CORE   coincide il solo nucleo numerico (copre i disallineamenti di suffisso
            lettera fra Excel ed ERP: Excel 226012 <-> ERP 226012K,
            Excel 226034B <-> ERP ...-226034)

   Tie-break fra candidati dello stesso livello:
     1. la parte con una riga in una commessa cliente (A_COM.COTYP = 'P'), piu' recente
     2. la parte non bloccata
     3. IDART piu' alto (la piu' recente)

   NOTA: la descrizione (PADSC) non e' utilizzabile come chiave, alterna
   DRYER/DRIER e SN/S.N. con spaziature irregolari. Nemmeno PATYP e' filtrabile:
   le macchine sono sia 'F' che 'S'.
   ============================================================================= */
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

IF OBJECT_ID('dbo.X_fn_trascodeSNpart', 'FN') IS NOT NULL
    DROP FUNCTION dbo.X_fn_trascodeSNpart;
GO
IF OBJECT_ID('dbo.X_fn_trascodeSNpart_Detail', 'IF') IS NOT NULL
    DROP FUNCTION dbo.X_fn_trascodeSNpart_Detail;
GO

/* ---------------------------------------------------------------------------
   Motore: inline table valued function.
   Restituisce 0 o 1 riga (PACOD, MatchLevel, Candidates, IDART).
   --------------------------------------------------------------------------- */
CREATE FUNCTION dbo.X_fn_trascodeSNpart_Detail (@SN varchar(30))
RETURNS TABLE
AS
RETURN
(
    WITH norm AS
    (
        /* maiuscolo, senza spazi ne' separatori */
        SELECT Norm = UPPER(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(
                        LTRIM(RTRIM(ISNULL(@SN, ''))), ' ', ''), '.', ''), '-', ''), '/', ''), '_', ''))
    ),
    sn AS
    (
        SELECT Norm,
               /* nucleo numerico = cifre iniziali (226019K -> 226019) */
               Core = CASE WHEN PATINDEX('%[^0-9]%', Norm) = 0 THEN Norm
                           ELSE LEFT(Norm, PATINDEX('%[^0-9]%', Norm) - 1) END
        FROM norm
        WHERE LEN(Norm) >= 4
    ),
    sn_ok AS
    (
        /* senza un nucleo numerico di almeno 4 cifre non si cerca nulla:
           eviterebbe il filtro sul LIKE e scansionerebbe tutta A_PAR */
        SELECT Norm, Core FROM sn WHERE LEN(Core) >= 4
    ),
    cand AS
    (
        SELECT p.PACOD,
               p.IDART,
               IsBlocked = ISNULL(p.IsBlocked, CAST(0 AS bit)),
               /* porzione del codice dopo l'ultimo trattino */
               SnPart = RIGHT(RTRIM(p.PACOD), CHARINDEX('-', REVERSE(RTRIM(p.PACOD))) - 1),
               sn.Norm,
               sn.Core
        FROM sn_ok AS sn
        INNER JOIN dbo.A_PAR AS p
                ON  p.PACOD LIKE '%-%'
                AND p.PACOD LIKE '%-' + sn.Core + '%'
    ),
    lvl AS
    (
        SELECT c.PACOD,
               c.IDART,
               c.IsBlocked,
               MatchLevel = CASE
                                WHEN c.SnPart = c.Norm THEN 1
                                WHEN CASE WHEN PATINDEX('%[^0-9]%', c.SnPart) = 0 THEN c.SnPart
                                          ELSE LEFT(c.SnPart, PATINDEX('%[^0-9]%', c.SnPart) - 1) END = c.Core
                                     AND LEN(c.SnPart) - LEN(c.Core) BETWEEN 0 AND 1
                                THEN 2
                            END
        FROM cand AS c
    ),
    scored AS
    (
        SELECT l.PACOD,
               l.IDART,
               l.IsBlocked,
               l.MatchLevel,
               CandidatesAtLevel = COUNT(*) OVER (PARTITION BY l.MatchLevel),
               HasOrderRow = CASE WHEN EXISTS (SELECT 1
                                               FROM dbo.L_CMPA AS lc
                                               INNER JOIN dbo.A_COM AS ac ON ac.CONUM = lc.CONUM
                                               WHERE lc.PACOD = l.PACOD AND ac.COTYP = 'P')
                                  THEN 1 ELSE 0 END,
               MaxConum    = ISNULL((SELECT MAX(lc.CONUM)
                                     FROM dbo.L_CMPA AS lc
                                     INNER JOIN dbo.A_COM AS ac ON ac.CONUM = lc.CONUM
                                     WHERE lc.PACOD = l.PACOD AND ac.COTYP = 'P'), 0)
        FROM lvl AS l
        WHERE l.MatchLevel IS NOT NULL
    )
    SELECT TOP (1)
           PACOD      = RTRIM(s.PACOD),
           MatchLevel = CASE s.MatchLevel WHEN 1 THEN 'EXACT' ELSE 'CORE' END,
           Candidates = s.CandidatesAtLevel,
           IDART      = s.IDART
    FROM scored AS s
    ORDER BY s.MatchLevel ASC,      -- EXACT prima di CORE
             s.HasOrderRow DESC,    -- chi ha una riga di commessa cliente
             s.MaxConum    DESC,    -- la commessa piu' recente
             s.IsBlocked   ASC,     -- non bloccata
             s.IDART       DESC     -- la parte piu' recente
);
GO

/* ---------------------------------------------------------------------------
   Wrapper scalare: restituisce il solo PACOD (NULL se non trascodificabile).
   --------------------------------------------------------------------------- */
CREATE FUNCTION dbo.X_fn_trascodeSNpart (@SN varchar(30))
RETURNS varchar(20)
AS
BEGIN
    DECLARE @PACOD varchar(20);
    SELECT @PACOD = PACOD FROM dbo.X_fn_trascodeSNpart_Detail(@SN);
    RETURN @PACOD;
END
GO

PRINT 'X_fn_trascodeSNpart / _Detail : create.';
GO
