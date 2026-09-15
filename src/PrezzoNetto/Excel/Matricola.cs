using System.Globalization;
using System.Text;

namespace PrezzoNetto.Excel;

/// <summary>
/// Normalizzazione della matricola.
///
/// Nel file convivono numeri veri (226035), testo con suffisso maiuscolo
/// (226019K), suffisso minuscolo (226134k) e valori con spazi accidentali.
/// La regola e' la stessa applicata da X_fn_trascodeSNpart lato SQL:
/// maiuscolo, niente spazi ne' separatori.
/// </summary>
public static class Matricola
{
    public static string? Normalize(object? raw)
    {
        if (raw is null) return null;

        string text = raw switch
        {
            double d => FormatNumber(d),
            decimal m => FormatNumber((double)m),
            int i => i.ToString(CultureInfo.InvariantCulture),
            long l => l.ToString(CultureInfo.InvariantCulture),
            _ => raw.ToString() ?? ""
        };

        var sb = new StringBuilder(text.Length);
        foreach (var c in text)
        {
            if (char.IsLetterOrDigit(c))
                sb.Append(char.ToUpperInvariant(c));
        }

        var result = sb.ToString();
        return result.Length == 0 ? null : result;
    }

    /// <summary>Una matricola letta come numero non deve diventare "226035.0".</summary>
    private static string FormatNumber(double d)
        => Math.Abs(d % 1) < 1e-9
            ? ((long)Math.Round(d)).ToString(CultureInfo.InvariantCulture)
            : d.ToString("0.####", CultureInfo.InvariantCulture);
}
