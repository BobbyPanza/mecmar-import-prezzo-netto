using System.ComponentModel;
using System.IO;
using System.Windows;
using System.Windows.Data;
using Microsoft.Win32;
using PrezzoNetto.Configuration;
using PrezzoNetto.Import;
using PrezzoNetto.Model;
using PrezzoNetto.Report;

namespace PrezzoNetto;

public partial class MainWindow : Window
{
    private readonly AppSettings _settings;
    private readonly ImportService _service;
    private ImportSession? _session;

    private static readonly string LastPathFile = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "Intesi", "PrezzoNetto", "lastfile.txt");

    public MainWindow(AppSettings settings)
    {
        InitializeComponent();

        _settings = settings;
        _service = new ImportService(settings);

        TxtSheet.Text = settings.Excel.SheetName;
        TxtFile.Text = LoadLastPath() ?? "";
    }

    // ------------------------------------------------------------ selezione

    private void BtnBrowse_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Title = "Seleziona il file delle quotazioni",
            Filter = "Cartelle di lavoro Excel (*.xlsx;*.xlsm)|*.xlsx;*.xlsm|Tutti i file (*.*)|*.*",
            CheckFileExists = true
        };

        var current = TxtFile.Text;
        if (!string.IsNullOrWhiteSpace(current))
        {
            var dir = Path.GetDirectoryName(current);
            if (Directory.Exists(dir)) dialog.InitialDirectory = dir;
            dialog.FileName = Path.GetFileName(current);
        }

        if (dialog.ShowDialog(this) == true)
            TxtFile.Text = dialog.FileName;
    }

    private void Window_DragOver(object sender, DragEventArgs e)
    {
        e.Effects = e.Data.GetDataPresent(DataFormats.FileDrop) ? DragDropEffects.Copy : DragDropEffects.None;
        e.Handled = true;
    }

    private void Window_Drop(object sender, DragEventArgs e)
    {
        if (e.Data.GetData(DataFormats.FileDrop) is string[] { Length: > 0 } files)
            TxtFile.Text = files[0];
    }

    // -------------------------------------------------------------- analisi

    private void BtnAnalyze_Click(object sender, RoutedEventArgs e)
    {
        var path = TxtFile.Text.Trim();
        if (string.IsNullOrWhiteSpace(path))
        {
            MessageBox.Show(this, "Selezionare il file da importare.", "Import Prezzo Netto",
                            MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        Run("Analisi in corso...", () =>
        {
            _session = _service.Analyze(path, TxtSheet.Text.Trim(), ChkDryRun.IsChecked == true);
            SaveLastPath(path);
            Bind();

            if (_session.Resumed)
                MessageBox.Show(this,
                    $"Ripresa dell'esecuzione {_session.RunId}, interrotta in precedenza sullo stesso file.\n" +
                    "Le righe gia' scritte in ERP non verranno rielaborate.",
                    "Ripartenza", MessageBoxButton.OK, MessageBoxImage.Information);
        });
    }

    // ---------------------------------------------------------- applicazione

    private void BtnApply_Click(object sender, RoutedEventArgs e)
    {
        if (_session is null) return;

        if (ChkDryRun.IsChecked == true)
        {
            MessageBox.Show(this, "La modalita' simulazione e' attiva: nessuna scrittura verra' eseguita.",
                            "Import Prezzo Netto", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var changes = _session.Changes;
        if (changes == 0)
        {
            MessageBox.Show(this, "Non ci sono righe da aggiornare.", "Import Prezzo Netto",
                            MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var confirm = MessageBox.Show(this,
            $"Verranno aggiornate {changes} righe di commessa e altrettante parti in Factory.\n\n" +
            "L'operazione e' tracciata e annullabile con X_sp_ImportPrezzoNetto_Rollback.\n\nProcedere?",
            "Conferma applicazione", MessageBoxButton.YesNo, MessageBoxImage.Question);

        if (confirm != MessageBoxResult.Yes) return;

        Run("Applicazione in corso...", () =>
        {
            _service.Apply(_session);
            Bind();
            MessageBox.Show(this, $"Aggiornamento completato: {changes} righe scritte in ERP.",
                            "Import Prezzo Netto", MessageBoxButton.OK, MessageBoxImage.Information);
        });
    }

    // ------------------------------------------------------------- report

    private void BtnExport_Click(object sender, RoutedEventArgs e)
    {
        if (_session is null) return;

        Run("Esportazione...", () =>
        {
            var folder = Path.IsPathRooted(_settings.Import.ReportFolder)
                ? _settings.Import.ReportFolder
                : Path.Combine(AppContext.BaseDirectory, _settings.Import.ReportFolder);

            var path = ReportWriter.Write(_session, folder);
            LblStatus.Text = $"Report salvato in {path}";

            if (MessageBox.Show(this, $"Report salvato in:\n{path}\n\nAprirlo adesso?", "Report",
                                MessageBoxButton.YesNo, MessageBoxImage.Information) == MessageBoxResult.Yes)
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(path) { UseShellExecute = true });
            }
        });
    }

    // ------------------------------------------------------------- griglia

    private void Bind()
    {
        if (_session is null) return;

        Grid.ItemsSource = _session.Rows;
        ApplyFilter();

        var s = _session;
        LblSummary.Text =
            $"Run {s.RunId}  |  righe {s.Rows.Count}  |  con prezzo {s.WithPrice}  |  " +
            $"trascodificate {s.Matched}  |  da aggiornare {s.Changes}  |  invariate {s.Unchanged}  |  " +
            $"warning {s.Warnings}  |  errori {s.Errors}";

        BtnApply.IsEnabled = s.Changes > 0 && ChkDryRun.IsChecked != true;
        BtnExport.IsEnabled = true;
        LblStatus.Text = $"Foglio \"{s.Source.SheetName}\", intestazioni alla riga {s.Source.HeaderRow}.";
    }

    private void Filter_Changed(object sender, RoutedEventArgs e) => ApplyFilter();

    private void ApplyFilter()
    {
        if (Grid.ItemsSource is null) return;

        var view = CollectionViewSource.GetDefaultView(Grid.ItemsSource);
        var onlyChanges = ChkOnlyChanges.IsChecked == true;
        var onlyProblems = ChkOnlyProblems.IsChecked == true;

        view.Filter = onlyChanges || onlyProblems
            ? o =>
            {
                var r = (ImportRow)o;
                if (onlyChanges && !r.IsChange) return false;
                if (onlyProblems && !r.IsError && !r.HasWarning) return false;
                return true;
            }
        : null;
    }

    // -------------------------------------------------------------- utilita'

    private void Run(string status, Action action)
    {
        System.Windows.Input.Mouse.OverrideCursor = System.Windows.Input.Cursors.Wait;
        LblStatus.Text = status;
        IsEnabled = false;

        try
        {
            action();
        }
        catch (Exception ex)
        {
            LblStatus.Text = "Operazione non riuscita.";
            MessageBox.Show(this, ex.Message, "Errore", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            IsEnabled = true;
            System.Windows.Input.Mouse.OverrideCursor = null;
        }
    }

    private static string? LoadLastPath()
    {
        try { return File.Exists(LastPathFile) ? File.ReadAllText(LastPathFile).Trim() : null; }
        catch { return null; }
    }

    private static void SaveLastPath(string path)
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(LastPathFile)!);
            File.WriteAllText(LastPathFile, path);
        }
        catch { /* preferenza non essenziale */ }
    }
}
