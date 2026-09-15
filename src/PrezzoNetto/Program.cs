using System.IO;
using System.Runtime.InteropServices;
using PrezzoNetto.Configuration;

namespace PrezzoNetto;

public static class Program
{
    public const int ExitOk = 0;
    public const int ExitWarning = 1;
    public const int ExitError = 2;

    [STAThread]
    public static int Main(string[] args)
    {
        CommandLineOptions options;
        try
        {
            options = CommandLineOptions.Parse(args);
        }
        catch (Exception ex)
        {
            AttachToParentConsole();
            Console.Error.WriteLine(ex.Message);
            Console.Error.WriteLine();
            Console.Error.WriteLine(CommandLineOptions.HelpText);
            return ExitError;
        }

        if (options.ShowHelp)
        {
            AttachToParentConsole();
            Console.WriteLine(CommandLineOptions.HelpText);
            return ExitOk;
        }

        AppSettings settings;
        try
        {
            settings = AppSettings.Load(AppContext.BaseDirectory);
        }
        catch (Exception ex)
        {
            return Fail(options, $"Configurazione non valida: {ex.Message}");
        }

        if (options.ConnectionString is not null) settings.ConnectionString = options.ConnectionString;
        if (options.ReportFolder is not null) settings.Import.ReportFolder = options.ReportFolder;
        if (options.Decimals is not null) settings.Import.Decimals = options.Decimals.Value;

        // percorso passato come argomento: esecuzione non presidiata
        if (options.Headless)
        {
            AttachToParentConsole();
            return ConsoleRunner.Run(settings, options);
        }

        var app = new App();
        app.InitializeComponent();
        app.Run(new MainWindow(settings));
        return ExitOk;
    }

    private static int Fail(CommandLineOptions options, string message)
    {
        if (options.Headless)
        {
            AttachToParentConsole();
            Console.Error.WriteLine(message);
        }
        else
        {
            System.Windows.MessageBox.Show(message, "Import Prezzo Netto",
                System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
        }
        return ExitError;
    }

    /// <summary>
    /// L'eseguibile e' una WinExe: senza questo aggancio l'output da riga di
    /// comando non comparirebbe nella console chiamante.
    /// </summary>
    private static void AttachToParentConsole()
    {
        if (!AttachConsole(-1)) return;

        var stdout = new StreamWriter(Console.OpenStandardOutput()) { AutoFlush = true };
        var stderr = new StreamWriter(Console.OpenStandardError()) { AutoFlush = true };
        Console.SetOut(stdout);
        Console.SetError(stderr);
    }

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool AttachConsole(int dwProcessId);
}
