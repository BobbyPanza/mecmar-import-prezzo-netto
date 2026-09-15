namespace PrezzoNetto.Configuration;

/// <summary>
/// Argomenti accettati dall'eseguibile.
/// Senza argomenti si apre l'interfaccia grafica; con un percorso si esegue
/// l'importazione senza presidio e si esce con un exit code.
/// </summary>
public sealed class CommandLineOptions
{
    public string? FilePath { get; private set; }
    public string? SheetName { get; private set; }
    public string? ConnectionString { get; private set; }
    public string? ReportFolder { get; private set; }
    public byte? Decimals { get; private set; }
    public bool DryRun { get; private set; }
    public bool Quiet { get; private set; }
    public bool ShowHelp { get; private set; }

    public bool Headless => FilePath is not null;

    public static CommandLineOptions Parse(string[] args)
    {
        var o = new CommandLineOptions();

        for (var i = 0; i < args.Length; i++)
        {
            var a = args[i];

            switch (a.ToLowerInvariant())
            {
                case "--help":
                case "-h":
                case "/?":
                    o.ShowHelp = true;
                    break;
                case "--dry-run":
                    o.DryRun = true;
                    break;
                case "--quiet":
                    o.Quiet = true;
                    break;
                case "--sheet":
                    o.SheetName = Next(args, ref i, a);
                    break;
                case "--conn":
                    o.ConnectionString = Next(args, ref i, a);
                    break;
                case "--report":
                    o.ReportFolder = Next(args, ref i, a);
                    break;
                case "--decimals":
                    o.Decimals = byte.Parse(Next(args, ref i, a));
                    break;
                default:
                    if (a.StartsWith('-'))
                        throw new ArgumentException($"Opzione sconosciuta: {a}");
                    if (o.FilePath is not null)
                        throw new ArgumentException("E' ammesso un solo percorso file.");
                    o.FilePath = a;
                    break;
            }
        }

        return o;
    }

    private static string Next(string[] args, ref int i, string option)
    {
        if (i + 1 >= args.Length)
            throw new ArgumentException($"Manca il valore per l'opzione {option}.");
        return args[++i];
    }

    public const string HelpText = """
        PrezzoNetto - importa il prezzo netto dal foglio "Listone Finale" in Factory.

          PrezzoNetto.exe                        apre l'interfaccia grafica
          PrezzoNetto.exe <file.xlsx> [opzioni]  esegue l'importazione e termina

        Opzioni:
          --sheet <nome>        foglio da leggere        (default: LISTONE FINALE)
          --dry-run             analizza senza scrivere su ERP
          --conn <stringa>      stringa di connessione alternativa
          --report <cartella>   cartella del report
          --decimals <n>        decimali di arrotondamento (default: 2)
          --quiet               solo il riepilogo finale
          --help                questo testo

        Exit code:
          0  completato senza problemi
          1  completato con warning (righe non trascodificate, ambigue, ...)
          2  errore bloccante
        """;
}
