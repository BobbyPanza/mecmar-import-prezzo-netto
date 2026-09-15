using System.Data;
using Microsoft.Data.SqlClient;
using PrezzoNetto.Model;

namespace PrezzoNetto.Data;

/// <summary>
/// Accesso al DB Factory.
///
/// L'applicativo non emette alcun UPDATE su L_CMPA o A_PAR: carica lo staging
/// e chiama le stored procedure X_sp_ImportPrezzoNetto_*. Tutta la logica di
/// aggiornamento resta modificabile in SQL.
/// </summary>
public sealed class ImportRepository
{
    private readonly string _connectionString;

    public ImportRepository(string connectionString) => _connectionString = connectionString;

    private SqlConnection Open()
    {
        var cn = new SqlConnection(_connectionString);
        cn.Open();
        return cn;
    }

    public void TestConnection()
    {
        using var cn = Open();
        using var cmd = new SqlCommand("SELECT 1", cn);
        cmd.ExecuteScalar();
    }

    // ------------------------------------------------------------- run

    public (int RunId, bool Resumed) BeginRun(ExcelReadResult source, string mode, bool resume)
    {
        using var cn = Open();
        using var cmd = new SqlCommand("dbo.X_sp_ImportPrezzoNetto_BeginRun", cn)
        {
            CommandType = CommandType.StoredProcedure
        };

        cmd.Parameters.AddWithValue("@SourceFile", source.FilePath);
        cmd.Parameters.AddWithValue("@SheetName", source.SheetName);
        cmd.Parameters.AddWithValue("@Mode", mode);
        cmd.Parameters.AddWithValue("@SourceFileHash", source.FileHash);
        cmd.Parameters.AddWithValue("@SourceFileModifiedAt", source.FileModifiedAt);
        cmd.Parameters.AddWithValue("@UserName", Environment.UserName);
        cmd.Parameters.AddWithValue("@MachineName", Environment.MachineName);
        cmd.Parameters.AddWithValue("@Resume", resume);

        using var reader = cmd.ExecuteReader();
        if (!reader.Read())
            throw new InvalidOperationException("BeginRun non ha restituito un RunId.");

        return (reader.GetInt32(0), reader.GetBoolean(1));
    }

    /// <summary>Carica le righe lette dall'Excel nello staging.</summary>
    public void LoadRows(int runId, IEnumerable<ExcelRowData> rows)
    {
        var table = BuildStagingTable(runId, rows);
        if (table.Rows.Count == 0) return;

        using var cn = Open();

        // il reload di un run gia' caricato riparte pulito
        using (var del = new SqlCommand(
                   "DELETE FROM dbo.X_ImportPrezzoNetto_Row WHERE RunId = @RunId AND ProcessStatus <> 'DONE'", cn))
        {
            del.Parameters.AddWithValue("@RunId", runId);
            del.ExecuteNonQuery();
        }

        // le righe gia' applicate non vanno ricaricate
        var applied = new HashSet<int>();
        using (var sel = new SqlCommand(
                   "SELECT ExcelRow FROM dbo.X_ImportPrezzoNetto_Row WHERE RunId = @RunId", cn))
        {
            sel.Parameters.AddWithValue("@RunId", runId);
            using var r = sel.ExecuteReader();
            while (r.Read()) applied.Add(r.GetInt32(0));
        }

        if (applied.Count > 0)
        {
            for (var i = table.Rows.Count - 1; i >= 0; i--)
                if (applied.Contains((int)table.Rows[i]["ExcelRow"]))
                    table.Rows.RemoveAt(i);
        }

        if (table.Rows.Count == 0) return;

        using var bulk = new SqlBulkCopy(cn) { DestinationTableName = "dbo.X_ImportPrezzoNetto_Row" };
        foreach (DataColumn c in table.Columns)
            bulk.ColumnMappings.Add(c.ColumnName, c.ColumnName);
        bulk.WriteToServer(table);
    }

    private static DataTable BuildStagingTable(int runId, IEnumerable<ExcelRowData> rows)
    {
        var t = new DataTable();
        t.Columns.Add("RunId", typeof(int));
        t.Columns.Add("ExcelRow", typeof(int));
        t.Columns.Add("MatricolaRaw", typeof(string));
        t.Columns.Add("Matricola", typeof(string));
        t.Columns.Add("PrezzoNetto", typeof(decimal));
        t.Columns.Add("CommessaExcel", typeof(int));
        t.Columns.Add("Modello", typeof(string));
        t.Columns.Add("Cliente", typeof(string));
        t.Columns.Add("DataConsegna", typeof(DateTime));

        foreach (var r in rows)
        {
            t.Rows.Add(
                runId,
                r.ExcelRow,
                (object?)r.MatricolaRaw ?? DBNull.Value,
                (object?)r.Matricola ?? DBNull.Value,
                (object?)r.PrezzoNetto ?? DBNull.Value,
                (object?)r.CommessaExcel ?? DBNull.Value,
                (object?)Truncate(r.Modello, 160) ?? DBNull.Value,
                (object?)Truncate(r.Cliente, 160) ?? DBNull.Value,
                (object?)r.DataConsegna ?? DBNull.Value);
        }

        return t;
    }

    private static string? Truncate(string? s, int max)
        => s is null || s.Length <= max ? s : s[..max];

    // --------------------------------------------------------- analisi

    public List<ImportRow> Analyze(int runId, byte decimals, decimal warnDeltaPct)
    {
        using var cn = Open();
        using var cmd = new SqlCommand("dbo.X_sp_ImportPrezzoNetto_Analyze", cn)
        {
            CommandType = CommandType.StoredProcedure,
            CommandTimeout = 300
        };
        cmd.Parameters.AddWithValue("@RunId", runId);
        cmd.Parameters.AddWithValue("@Decimals", decimals);
        cmd.Parameters.AddWithValue("@WarnDeltaPct", warnDeltaPct);

        using var reader = cmd.ExecuteReader();
        return ReadRows(reader);
    }

    // ------------------------------------------------------ applicazione

    public List<ImportRow> Apply(int runId, bool dryRun)
    {
        using var cn = Open();
        using var cmd = new SqlCommand("dbo.X_sp_ImportPrezzoNetto_Apply", cn)
        {
            CommandType = CommandType.StoredProcedure,
            CommandTimeout = 300
        };
        cmd.Parameters.AddWithValue("@RunId", runId);
        cmd.Parameters.AddWithValue("@DryRun", dryRun);
        cmd.Parameters.AddWithValue("@AppliedBy", Environment.UserName);

        using var reader = cmd.ExecuteReader();
        return dryRun ? new List<ImportRow>() : ReadRows(reader);
    }

    public List<ImportRow> GetRows(int runId)
    {
        using var cn = Open();
        using var cmd = new SqlCommand("dbo.X_sp_ImportPrezzoNetto_GetRows", cn)
        {
            CommandType = CommandType.StoredProcedure
        };
        cmd.Parameters.AddWithValue("@RunId", runId);

        using var reader = cmd.ExecuteReader();
        return ReadRows(reader);
    }

    public RunSummary GetSummary(int runId)
    {
        using var cn = Open();
        using var cmd = new SqlCommand(
            """
            SELECT RunId, RowsRead, RowsWithPrice, RowsMatched, RowsUpdated,
                   RowsUnchanged, RowsWarning, RowsError, Status
            FROM dbo.X_ImportPrezzoNetto_Run WHERE RunId = @RunId
            """, cn);
        cmd.Parameters.AddWithValue("@RunId", runId);

        using var r = cmd.ExecuteReader();
        if (!r.Read()) throw new InvalidOperationException($"Run {runId} non trovato.");

        return new RunSummary
        {
            RunId = r.GetInt32(0),
            RowsRead = r.GetInt32(1),
            RowsWithPrice = r.GetInt32(2),
            RowsMatched = r.GetInt32(3),
            RowsUpdated = r.GetInt32(4),
            RowsUnchanged = r.GetInt32(5),
            RowsWarning = r.GetInt32(6),
            RowsError = r.GetInt32(7),
            Status = r.GetString(8)
        };
    }

    public void CloseRun(int runId, string status, string? error = null)
    {
        using var cn = Open();
        using var cmd = new SqlCommand(
            """
            UPDATE dbo.X_ImportPrezzoNetto_Run
            SET Status = @Status, FinishedAt = SYSDATETIME(),
                ErrorMessage = COALESCE(@Error, ErrorMessage)
            WHERE RunId = @RunId AND Status = 'RUNNING'
            """, cn);
        cmd.Parameters.AddWithValue("@RunId", runId);
        cmd.Parameters.AddWithValue("@Status", status);
        cmd.Parameters.AddWithValue("@Error", (object?)error ?? DBNull.Value);
        cmd.ExecuteNonQuery();
    }

    // ------------------------------------------------------------ lettura

    private static List<ImportRow> ReadRows(SqlDataReader reader)
    {
        var list = new List<ImportRow>();

        // la Apply restituisce prima eventuali result set di servizio
        do
        {
            if (!HasColumn(reader, "ExcelRow") || !HasColumn(reader, "Action")) continue;

            while (reader.Read())
            {
                list.Add(new ImportRow
                {
                    ExcelRow = Get<int>(reader, "ExcelRow"),
                    MatricolaRaw = Get<string>(reader, "MatricolaRaw"),
                    Matricola = Get<string>(reader, "Matricola"),
                    Modello = Get<string>(reader, "Modello"),
                    Cliente = Get<string>(reader, "Cliente"),
                    DataConsegna = Get<DateTime?>(reader, "DataConsegna"),
                    CommessaExcel = Get<int?>(reader, "CommessaExcel"),
                    PACOD = Get<string>(reader, "PACOD"),
                    MatchLevel = Get<string>(reader, "MatchLevel"),
                    Candidates = Get<int?>(reader, "Candidates"),
                    COCOD = Get<string>(reader, "COCOD"),
                    CONUM = Get<int?>(reader, "CONUM"),
                    IDROW = Get<int?>(reader, "IDROW"),
                    PrezzoErp = Get<decimal?>(reader, "PrezzoErp"),
                    PrezzoNetto = Get<decimal?>(reader, "PrezzoNetto"),
                    Delta = Get<decimal?>(reader, "Delta"),
                    DeltaPct = Get<decimal?>(reader, "DeltaPct"),
                    PrezzoPrecedenteImport = Get<decimal?>(reader, "PrezzoPrecedenteImport"),
                    Action = Get<string>(reader, "Action"),
                    ProcessStatus = Get<string>(reader, "ProcessStatus"),
                    ProcessAttempts = Get<int>(reader, "ProcessAttempts"),
                    ProcessedAt = Get<DateTime?>(reader, "ProcessedAt"),
                    Applied = Get<bool>(reader, "Applied"),
                    HasWarning = Get<bool>(reader, "HasWarning"),
                    Message = Get<string>(reader, "Message")
                });
            }

            if (list.Count > 0) break;
        }
        while (reader.NextResult());

        return list.OrderBy(r => r.ExcelRow).ToList();
    }

    private static bool HasColumn(SqlDataReader reader, string name)
    {
        for (var i = 0; i < reader.FieldCount; i++)
            if (string.Equals(reader.GetName(i), name, StringComparison.OrdinalIgnoreCase))
                return true;
        return false;
    }

    private static T? Get<T>(SqlDataReader reader, string name)
    {
        if (!HasColumn(reader, name)) return default;
        var i = reader.GetOrdinal(name);
        if (reader.IsDBNull(i)) return default;
        return (T)Convert.ChangeType(reader.GetValue(i), Nullable.GetUnderlyingType(typeof(T)) ?? typeof(T));
    }
}
