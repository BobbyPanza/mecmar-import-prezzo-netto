namespace PrezzoNetto.Model;

/// <summary>Una riga letta dal foglio, gia' normalizzata.</summary>
public sealed class ExcelRowData
{
    public int ExcelRow { get; init; }
    public string? MatricolaRaw { get; init; }
    public string? Matricola { get; init; }
    public decimal? PrezzoNetto { get; init; }
    public int? CommessaExcel { get; init; }
    public string? Modello { get; init; }
    public string? Cliente { get; init; }
    public DateTime? DataConsegna { get; init; }
}

/// <summary>Esito della lettura del foglio.</summary>
public sealed class ExcelReadResult
{
    public required string FilePath { get; init; }
    public required string SheetName { get; init; }
    public required int HeaderRow { get; init; }
    public required IReadOnlyList<ExcelRowData> Rows { get; init; }
    public required string FileHash { get; init; }
    public required DateTime FileModifiedAt { get; init; }
    public List<string> Warnings { get; } = new();
}
