using System.Globalization;
using System.IO;
using System.Security.Cryptography;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using PrezzoNetto.Configuration;
using PrezzoNetto.Model;

namespace PrezzoNetto.Excel;

/// <summary>
/// Lettura del foglio "Listone Finale".
///
/// Due accortezze imposte dal file reale:
///  - la riga di intestazione non e' la prima (nel file campione e' la 2:
///    la riga 1 contiene i codici opzione), quindi le colonne si cercano
///    per nome scandendo le prime righe;
///  - la colonna del prezzo netto contiene sia formule sia valori digitati
///    a mano, quindi delle formule si legge il valore memorizzato e non si
///    ricalcola nulla.
/// </summary>
public sealed class ListoneReader
{
    private readonly ExcelSettings _settings;

    public ListoneReader(ExcelSettings settings) => _settings = settings;

    public ExcelReadResult Read(string filePath, string? sheetOverride = null)
    {
        if (!File.Exists(filePath))
            throw new FileNotFoundException($"File non trovato: {filePath}", filePath);

        var sheetName = sheetOverride ?? _settings.SheetName;
        var fileInfo = new FileInfo(filePath);
        var hash = ComputeHash(filePath);

        // FileShare.ReadWrite: il file deve essere leggibile anche se aperto in Excel.
        using var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read,
                                          FileShare.ReadWrite | FileShare.Delete);
        using var workbook = new XLWorkbook(stream);

        var sheet = FindSheet(workbook, sheetName)
            ?? throw new InvalidOperationException(
                $"Foglio \"{sheetName}\" non trovato. Fogli presenti: " +
                string.Join(", ", workbook.Worksheets.Select(w => $"\"{w.Name}\"")));

        var (headerRow, columns) = FindHeader(sheet);

        var rows = new List<ExcelRowData>();
        var blanks = 0;
        var lastRow = sheet.LastRowUsed()?.RowNumber() ?? headerRow;

        for (var r = headerRow + 1; r <= lastRow; r++)
        {
            var matricolaRaw = ReadCell(sheet, r, columns.Matricola);
            var prezzoRaw = ReadCell(sheet, r, columns.PrezzoNetto);

            if (matricolaRaw is null && prezzoRaw is null)
            {
                if (++blanks >= _settings.BlankRowsToStop) break;
                continue;
            }
            blanks = 0;

            var matricola = Matricola.Normalize(matricolaRaw);
            if (matricola is null) continue;   // riga senza matricola: ignorata

            rows.Add(new ExcelRowData
            {
                ExcelRow = r,
                MatricolaRaw = matricolaRaw?.ToString(),
                Matricola = matricola,
                PrezzoNetto = ToDecimal(prezzoRaw),
                CommessaExcel = ToInt(ReadCell(sheet, r, columns.Commessa)),
                Modello = ToText(ReadCell(sheet, r, columns.Modello)),
                Cliente = ToText(ReadCell(sheet, r, columns.Cliente)),
                DataConsegna = ToDate(ReadCell(sheet, r, columns.DataConsegna))
            });
        }

        var result = new ExcelReadResult
        {
            FilePath = filePath,
            SheetName = sheet.Name,
            HeaderRow = headerRow,
            Rows = rows,
            FileHash = hash,
            FileModifiedAt = fileInfo.LastWriteTime
        };

        if (rows.Count == 0)
            result.Warnings.Add($"Nessuna riga con matricola trovata sotto la riga {headerRow}.");

        return result;
    }

    // ---------------------------------------------------------------- foglio

    private static IXLWorksheet? FindSheet(XLWorkbook workbook, string name)
    {
        var target = NormalizeHeader(name);
        return workbook.Worksheets.FirstOrDefault(w => NormalizeHeader(w.Name) == target);
    }

    // ---------------------------------------------------------- intestazioni

    private sealed record ColumnMap(int Matricola, int PrezzoNetto,
                                    int? Commessa, int? Modello, int? Cliente, int? DataConsegna);

    private (int HeaderRow, ColumnMap Columns) FindHeader(IXLWorksheet sheet)
    {
        var wanted = _settings.Columns;
        var lastCol = sheet.LastColumnUsed()?.ColumnNumber() ?? 1;
        var maxRow = Math.Min(_settings.HeaderSearchRows, sheet.LastRowUsed()?.RowNumber() ?? 1);

        var seen = new List<string>();

        for (var r = 1; r <= maxRow; r++)
        {
            var map = new Dictionary<string, int>(StringComparer.Ordinal);
            for (var c = 1; c <= lastCol; c++)
            {
                var text = NormalizeHeader(sheet.Cell(r, c).GetFormattedString());
                if (text.Length == 0) continue;
                map.TryAdd(text, c);
            }

            var matricola = Find(map, wanted.Matricola);
            var prezzo = Find(map, wanted.PrezzoNetto);

            if (matricola is not null && prezzo is not null)
            {
                return (r, new ColumnMap(
                    matricola.Value,
                    prezzo.Value,
                    Find(map, wanted.Commessa),
                    Find(map, wanted.Modello),
                    Find(map, wanted.Cliente),
                    Find(map, wanted.DataConsegna)));
            }

            if (map.Count > 0) seen.AddRange(map.Keys);
        }

        throw new InvalidOperationException(
            $"Nelle prime {maxRow} righe del foglio \"{sheet.Name}\" non sono state trovate " +
            $"entrambe le colonne \"{wanted.Matricola}\" e \"{wanted.PrezzoNetto}\". " +
            $"Intestazioni lette: {string.Join(" | ", seen.Distinct().Take(60))}");
    }

    private static int? Find(Dictionary<string, int> map, string? header)
    {
        if (string.IsNullOrWhiteSpace(header)) return null;
        return map.TryGetValue(NormalizeHeader(header), out var col) ? col : null;
    }

    /// <summary>Maiuscolo, senza accenti di spaziatura, spazi multipli compattati.</summary>
    private static string NormalizeHeader(string? s)
        => s is null ? "" : Regex.Replace(s.Trim(), @"\s+", " ").ToUpperInvariant();

    // ------------------------------------------------------------- celle

    /// <summary>
    /// Delle celle con formula si legge il valore memorizzato nel file:
    /// ClosedXML tenterebbe di rivalutare la formula e fallirebbe su quelle
    /// che referenziano altri fogli.
    /// </summary>
    private static object? ReadCell(IXLWorksheet sheet, int row, int? col)
    {
        if (col is null) return null;

        var cell = sheet.Cell(row, col.Value);
        XLCellValue value;

        if (cell.HasFormula)
        {
            try { value = cell.CachedValue; }
            catch { return null; }
        }
        else
        {
            value = cell.Value;
        }

        if (value.IsBlank || value.IsError) return null;

        return value.Type switch
        {
            XLDataType.Number => value.GetNumber(),
            XLDataType.Text => value.GetText(),
            XLDataType.DateTime => value.GetDateTime(),
            XLDataType.Boolean => value.GetBoolean(),
            XLDataType.TimeSpan => value.GetTimeSpan(),
            _ => null
        };
    }

    private static decimal? ToDecimal(object? v) => v switch
    {
        null => null,
        double d => (decimal)d,
        decimal m => m,
        int i => i,
        string s when decimal.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var p) => p,
        string s when decimal.TryParse(s, NumberStyles.Any, CultureInfo.CurrentCulture, out var p) => p,
        _ => null
    };

    private static int? ToInt(object? v)
    {
        var d = ToDecimal(v);
        if (d is null) return null;
        var rounded = Math.Round(d.Value);
        return rounded is >= int.MinValue and <= int.MaxValue ? (int)rounded : null;
    }

    private static string? ToText(object? v)
    {
        var s = v switch
        {
            null => null,
            double d => FormatIfWhole(d),
            DateTime dt => dt.ToString("yyyy-MM-dd"),
            _ => v.ToString()
        };
        s = s?.Trim();
        return string.IsNullOrEmpty(s) ? null : s;
    }

    private static string FormatIfWhole(double d)
        => Math.Abs(d % 1) < 1e-9
            ? ((long)Math.Round(d)).ToString(CultureInfo.InvariantCulture)
            : d.ToString(CultureInfo.InvariantCulture);

    private static DateTime? ToDate(object? v) => v switch
    {
        DateTime dt => dt,
        double d when d > 0 && d < 2958466 => DateTime.FromOADate(d),
        string s when DateTime.TryParse(s, out var p) => p,
        _ => null
    };

    // -------------------------------------------------------------- hash

    private static string ComputeHash(string path)
    {
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read,
                                          FileShare.ReadWrite | FileShare.Delete);
        using var sha = SHA256.Create();
        return Convert.ToHexString(sha.ComputeHash(stream)).ToLowerInvariant();
    }
}
