namespace PrezzoNetto.Model;

/// <summary>
/// Riga di staging cosi' come la restituisce X_sp_ImportPrezzoNetto_GetRows:
/// e' quello che si vede nella griglia e nel report.
/// </summary>
public sealed class ImportRow
{
    public int ExcelRow { get; set; }
    public string? MatricolaRaw { get; set; }
    public string? Matricola { get; set; }
    public string? Modello { get; set; }
    public string? Cliente { get; set; }
    public DateTime? DataConsegna { get; set; }
    public int? CommessaExcel { get; set; }
    public string? PACOD { get; set; }
    public string? MatchLevel { get; set; }
    public int? Candidates { get; set; }
    public string? COCOD { get; set; }
    public int? CONUM { get; set; }
    public int? IDROW { get; set; }
    public decimal? PrezzoErp { get; set; }
    public decimal? PrezzoNetto { get; set; }
    public decimal? Delta { get; set; }
    public decimal? DeltaPct { get; set; }
    public decimal? PrezzoPrecedenteImport { get; set; }
    public string? Action { get; set; }
    public string? ProcessStatus { get; set; }
    public int ProcessAttempts { get; set; }
    public DateTime? ProcessedAt { get; set; }
    public bool Applied { get; set; }
    public bool HasWarning { get; set; }
    public string? Message { get; set; }

    public bool IsError => ProcessStatus == "ERROR";
    public bool IsChange => Action is "NUOVO" or "MODIFICATO" or "RIALLINEATO";
}

/// <summary>Contatori di riepilogo di un'esecuzione.</summary>
public sealed class RunSummary
{
    public int RunId { get; set; }
    public int RowsRead { get; set; }
    public int RowsWithPrice { get; set; }
    public int RowsMatched { get; set; }
    public int RowsUpdated { get; set; }
    public int RowsUnchanged { get; set; }
    public int RowsWarning { get; set; }
    public int RowsError { get; set; }
    public string Status { get; set; } = "";
}
