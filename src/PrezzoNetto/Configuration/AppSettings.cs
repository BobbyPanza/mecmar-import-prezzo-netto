using Microsoft.Extensions.Configuration;

namespace PrezzoNetto.Configuration;

public sealed class AppSettings
{
    public string ConnectionString { get; set; } = "";
    public ExcelSettings Excel { get; set; } = new();
    public ImportSettings Import { get; set; } = new();

    public static AppSettings Load(string baseDirectory)
    {
        // appsettings.local.json sovrascrive il file base ed e' escluso dal
        // controllo di versione: e' li' che vanno le credenziali reali.
        var config = new ConfigurationBuilder()
            .SetBasePath(baseDirectory)
            .AddJsonFile("appsettings.json", optional: false, reloadOnChange: false)
            .AddJsonFile("appsettings.local.json", optional: true, reloadOnChange: false)
            .Build();

        var settings = new AppSettings();
        config.Bind(settings);
        return settings;
    }
}

public sealed class ExcelSettings
{
    public string SheetName { get; set; } = "LISTONE FINALE";

    /// <summary>Entro quante righe dall'alto cercare la riga di intestazione.</summary>
    public int HeaderSearchRows { get; set; } = 10;

    /// <summary>Quante righe vuote consecutive indicano la fine dei dati.</summary>
    public int BlankRowsToStop { get; set; } = 20;

    public ColumnSettings Columns { get; set; } = new();
}

public sealed class ColumnSettings
{
    public string Matricola { get; set; } = "MATRICOLA";
    public string PrezzoNetto { get; set; } = "PREZZO NETTO";
    public string? Commessa { get; set; } = "COMMESSA";
    public string? Modello { get; set; } = "MODEL";
    public string? Cliente { get; set; } = "CLIENTE";
    public string? DataConsegna { get; set; } = "DATA CONSEGNA";
}

public sealed class ImportSettings
{
    public byte Decimals { get; set; } = 2;
    public decimal WarnDeltaPct { get; set; } = 30m;
    public string ReportFolder { get; set; } = "Report";
}
