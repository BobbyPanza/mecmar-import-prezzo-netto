using System.IO;
using PrezzoNetto.Configuration;
using PrezzoNetto.Import;
using PrezzoNetto.Report;

namespace PrezzoNetto;

/// <summary>Esecuzione non presidiata: nessuna finestra, riepilogo ed exit code.</summary>
public static class ConsoleRunner
{
    public static int Run(AppSettings settings, CommandLineOptions options)
    {
        var service = new ImportService(settings);
        var quiet = options.Quiet;

        void Log(string message)
        {
            if (!quiet) Console.WriteLine(message);
        }

        ImportSession session;

        try
        {
            Log($"File   : {options.FilePath}");
            session = service.Analyze(options.FilePath!, options.SheetName, options.DryRun, options.Decimals);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"ERRORE: {ex.Message}");
            return Program.ExitError;
        }

        Log($"Foglio : {session.Source.SheetName} (intestazioni alla riga {session.Source.HeaderRow})");
        Log($"Run    : {session.RunId}{(session.Resumed ? " (ripresa di un'esecuzione interrotta)" : "")}");
        Log("");

        try
        {
            if (options.DryRun)
            {
                Log("Simulazione: nessuna scrittura eseguita.");
                service.CloseWithoutApplying(session);
            }
            else if (session.Changes > 0)
            {
                service.Apply(session);
                Log($"Applicate {session.Rows.Count(r => r.Applied)} righe.");
            }
            else
            {
                Log("Nessuna riga da aggiornare.");
                service.CloseWithoutApplying(session);
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"ERRORE in applicazione: {ex.Message}");
            return Program.ExitError;
        }

        // report
        try
        {
            var folder = Path.IsPathRooted(settings.Import.ReportFolder)
                ? settings.Import.ReportFolder
                : Path.Combine(AppContext.BaseDirectory, settings.Import.ReportFolder);

            var reportPath = ReportWriter.Write(session, folder);
            Log($"Report : {reportPath}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Report non scritto: {ex.Message}");
        }

        Console.WriteLine(
            $"Righe {session.Rows.Count} | con prezzo {session.WithPrice} | trascodificate {session.Matched} | " +
            $"da aggiornare {session.Changes} | applicate {session.Rows.Count(r => r.Applied)} | " +
            $"invariate {session.Unchanged} | warning {session.Warnings} | errori {session.Errors}");

        if (session.Errors > 0 && !quiet)
        {
            Console.WriteLine();
            Console.WriteLine("Righe non elaborate:");
            foreach (var r in session.Rows.Where(r => r.IsError))
                Console.WriteLine($"  riga {r.ExcelRow,4}  {r.Matricola,-10} {r.Action,-22} {r.Message}");
        }

        return session.Errors > 0 || session.Warnings > 0 ? Program.ExitWarning : Program.ExitOk;
    }
}
