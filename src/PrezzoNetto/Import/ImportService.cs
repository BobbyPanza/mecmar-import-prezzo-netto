using PrezzoNetto.Configuration;
using PrezzoNetto.Data;
using PrezzoNetto.Excel;
using PrezzoNetto.Model;

namespace PrezzoNetto.Import;

public sealed class ImportSession
{
    public required int RunId { get; init; }
    public required ExcelReadResult Source { get; init; }
    public required List<ImportRow> Rows { get; set; }
    public required bool Resumed { get; init; }

    public int Changes => Rows.Count(r => r.IsChange);
    public int Errors => Rows.Count(r => r.IsError);
    public int Warnings => Rows.Count(r => r.HasWarning);
    public int Unchanged => Rows.Count(r => r.Action == "INVARIATO");
    public int WithPrice => Rows.Count(r => r.PrezzoNetto > 0);
    public int Matched => Rows.Count(r => !string.IsNullOrEmpty(r.PACOD));
}

/// <summary>
/// Orchestrazione: legge il foglio, carica lo staging, lancia l'analisi
/// e, su richiesta, l'applicazione. Nessuna logica di aggiornamento qui.
/// </summary>
public sealed class ImportService
{
    private readonly AppSettings _settings;
    private readonly ImportRepository _repository;

    public ImportService(AppSettings settings)
    {
        _settings = settings;
        _repository = new ImportRepository(settings.ConnectionString);
    }

    public ImportRepository Repository => _repository;

    public ImportSession Analyze(string filePath, string? sheetOverride, bool dryRun, byte? decimalsOverride = null)
    {
        var reader = new ListoneReader(_settings.Excel);
        var source = reader.Read(filePath, sheetOverride);

        var (runId, resumed) = _repository.BeginRun(source, dryRun ? "DRYRUN" : "APPLY", resume: true);
        _repository.LoadRows(runId, source.Rows);

        var rows = _repository.Analyze(
            runId,
            decimalsOverride ?? _settings.Import.Decimals,
            _settings.Import.WarnDeltaPct);

        return new ImportSession { RunId = runId, Source = source, Rows = rows, Resumed = resumed };
    }

    public void Apply(ImportSession session)
    {
        session.Rows = _repository.Apply(session.RunId, dryRun: false);
    }

    public void CloseWithoutApplying(ImportSession session)
        => _repository.CloseRun(session.RunId, "COMPLETED");

    public RunSummary Summary(int runId) => _repository.GetSummary(runId);
}
