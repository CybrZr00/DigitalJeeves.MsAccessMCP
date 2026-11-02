using System.Data;
using System.Data.OleDb;

namespace DigitalJeeves.MsAccessMCP.Core;

/// <summary>
/// Executes OLE DB operations against an Access database for MCP tools.
///
/// Every public method funnels its actual work through <see cref="OleDbStaThread"/> — ACE/Jet is
/// apartment-threaded COM, and plain async/await resumes on arbitrary thread-pool (MTA) threads, which can
/// corrupt native COM state. This was previously masked when this ran only as an isolated subprocess
/// handling one sequential tool call at a time; hosting DigitalJeeves.MsAccessMCP.Core in-process removes that accidental
/// protection, so this now serializes explicitly. See OleDbStaThread's own doc comment for the full
/// rationale.
/// </summary>
internal sealed class AccessOleDb(AccessMcpOptions options)
{
    private readonly AccessMcpOptions _options = options;

    /// <summary>
    /// Creates an OLE DB connection using the configured provider.
    /// </summary>
    public OleDbConnection CreateConnection(string databasePath)
    {
        var builder = new OleDbConnectionStringBuilder
        {
            Provider = _options.OleDbProvider,
            DataSource = databasePath
        };

        return new OleDbConnection(builder.ConnectionString);
    }

    /// <summary>
    /// Lists user tables in the Access database.
    /// </summary>
    public Task<IReadOnlyList<string>> ListTablesAsync(string databasePath, CancellationToken cancellationToken) =>
        Task.Run(() => OleDbStaThread.Execute(() =>
        {
            using var connection = CreateConnection(databasePath);
            connection.Open();

            var tableNames = new List<string>();

            var schema = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, [null, null, null, "TABLE"]);
            if (schema is null)
            {
                return (IReadOnlyList<string>)tableNames;
            }

            foreach (DataRow row in schema.Rows)
            {
                if (row["TABLE_NAME"] is string name && !string.IsNullOrWhiteSpace(name))
                {
                    tableNames.Add(name);
                }
            }

            tableNames.Sort(StringComparer.OrdinalIgnoreCase);
            return (IReadOnlyList<string>)tableNames;
        }), cancellationToken);

    /// <summary>
    /// Lists views in the Access database.
    /// </summary>
    public Task<IReadOnlyList<string>> ListViewsAsync(string databasePath, CancellationToken cancellationToken) =>
        Task.Run(() => OleDbStaThread.Execute(() =>
        {
            using var connection = CreateConnection(databasePath);
            connection.Open();

            var viewNames = new List<string>();

            var schema = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, [null, null, null, "VIEW"]);
            if (schema is null)
            {
                return (IReadOnlyList<string>)viewNames;
            }

            foreach (DataRow row in schema.Rows)
            {
                if (row["TABLE_NAME"] is string name && !string.IsNullOrWhiteSpace(name))
                {
                    viewNames.Add(name);
                }
            }

            viewNames.Sort(StringComparer.OrdinalIgnoreCase);
            return (IReadOnlyList<string>)viewNames;
        }), cancellationToken);

    /// <summary>
    /// Lists stored queries in the Access database.
    /// </summary>
    public Task<IReadOnlyList<string>> ListQueriesAsync(string databasePath, CancellationToken cancellationToken) =>
        Task.Run(() => OleDbStaThread.Execute(() =>
        {
            using var connection = CreateConnection(databasePath);
            connection.Open();

            var queryNames = new List<string>();

            var schema = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Procedures, null);
            if (schema is null)
            {
                return (IReadOnlyList<string>)queryNames;
            }

            foreach (DataRow row in schema.Rows)
            {
                if (row["PROCEDURE_NAME"] is string name && !string.IsNullOrWhiteSpace(name))
                {
                    queryNames.Add(name);
                }
            }

            queryNames.Sort(StringComparer.OrdinalIgnoreCase);
            return (IReadOnlyList<string>)queryNames;
        }), cancellationToken);

    /// <summary>
    /// Lists indexes defined in the Access database.
    /// </summary>
    public Task<IReadOnlyList<IndexInfo>> ListIndexesAsync(string databasePath, string? tableName, CancellationToken cancellationToken) =>
        Task.Run(() => OleDbStaThread.Execute(() =>
        {
            using var connection = CreateConnection(databasePath);
            connection.Open();

            var indexes = new List<IndexInfo>();
            var restrictions = new object?[] { null, null, tableName, null, null };

            var schema = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Indexes, restrictions);
            if (schema is null)
            {
                return (IReadOnlyList<IndexInfo>)indexes;
            }

            foreach (DataRow row in schema.Rows)
            {
                var name = Convert.ToString(row["INDEX_NAME"]) ?? string.Empty;
                var table = Convert.ToString(row["TABLE_NAME"]) ?? string.Empty;

                if (string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(table))
                {
                    continue;
                }

                indexes.Add(new IndexInfo(
                    Name: name,
                    Table: table,
                    IsPrimaryKey: TryGetBool(row, "PRIMARY_KEY"),
                    IsUnique: TryGetBool(row, "UNIQUE"),
                    Type: Convert.ToString(row["INDEX_TYPE"]) ?? string.Empty));
            }

            return (IReadOnlyList<IndexInfo>)indexes
                .OrderBy(i => i.Table, StringComparer.OrdinalIgnoreCase)
                .ThenBy(i => i.Name, StringComparer.OrdinalIgnoreCase)
                .ToArray();

            static bool TryGetBool(DataRow row, string name)
            {
                if (!row.Table.Columns.Contains(name) || row[name] is DBNull)
                {
                    return false;
                }

                var value = row[name];
                return value switch
                {
                    bool b => b,
                    int i => i != 0,
                    short s => s != 0,
                    string str => str.Equals("true", StringComparison.OrdinalIgnoreCase) || str.Equals("yes", StringComparison.OrdinalIgnoreCase),
                    _ => Convert.ToBoolean(value)
                };
            }
        }), cancellationToken);

    /// <summary>
    /// Lists foreign key relationships defined in the Access database.
    /// </summary>
    public Task<IReadOnlyList<RelationshipInfo>> ListRelationshipsAsync(string databasePath, string? tableName, CancellationToken cancellationToken) =>
        Task.Run(() => OleDbStaThread.Execute(() =>
        {
            using var connection = CreateConnection(databasePath);
            connection.Open();

            var relationships = new List<RelationshipInfo>();
            var schema = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Foreign_Keys, null);
            if (schema is null)
            {
                return (IReadOnlyList<RelationshipInfo>)relationships;
            }

            foreach (DataRow row in schema.Rows)
            {
                var pkTable = Convert.ToString(row["PK_TABLE_NAME"]) ?? string.Empty;
                var pkColumn = Convert.ToString(row["PK_COLUMN_NAME"]) ?? string.Empty;
                var fkTable = Convert.ToString(row["FK_TABLE_NAME"]) ?? string.Empty;
                var fkColumn = Convert.ToString(row["FK_COLUMN_NAME"]) ?? string.Empty;

                if (string.IsNullOrWhiteSpace(pkTable) || string.IsNullOrWhiteSpace(fkTable))
                {
                    continue;
                }

                if (!string.IsNullOrWhiteSpace(tableName)
                    && !pkTable.Equals(tableName, StringComparison.OrdinalIgnoreCase)
                    && !fkTable.Equals(tableName, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                relationships.Add(new RelationshipInfo(
                    Name: Convert.ToString(row["FK_NAME"]) ?? string.Empty,
                    PrimaryTable: pkTable,
                    PrimaryColumn: pkColumn,
                    ForeignTable: fkTable,
                    ForeignColumn: fkColumn,
                    UpdateRule: Convert.ToString(row["UPDATE_RULE"]) ?? string.Empty,
                    DeleteRule: Convert.ToString(row["DELETE_RULE"]) ?? string.Empty));
            }

            return (IReadOnlyList<RelationshipInfo>)relationships
                .OrderBy(r => r.PrimaryTable, StringComparer.OrdinalIgnoreCase)
                .ThenBy(r => r.ForeignTable, StringComparer.OrdinalIgnoreCase)
                .ThenBy(r => r.Name, StringComparer.OrdinalIgnoreCase)
                .ToArray();
        }), cancellationToken);

    /// <summary>
    /// Lists columns for a table in the Access database.
    /// </summary>
    public Task<IReadOnlyList<ColumnInfo>> ListColumnsAsync(string databasePath, string tableName, CancellationToken cancellationToken)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(tableName);

        return Task.Run(() => OleDbStaThread.Execute(() =>
        {
            using var connection = CreateConnection(databasePath);
            connection.Open();

            var columns = new List<ColumnInfo>();

            var schema = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, [null, null, tableName, null]);
            if (schema is null)
            {
                return (IReadOnlyList<ColumnInfo>)columns;
            }

            foreach (DataRow row in schema.Rows)
            {
                columns.Add(new ColumnInfo(
                    Name: Convert.ToString(row["COLUMN_NAME"]) ?? string.Empty,
                    Ordinal: TryGetInt(row, "ORDINAL_POSITION"),
                    DataType: Convert.ToString(row["DATA_TYPE"]) ?? string.Empty,
                    IsNullable: TryGetBool(row, "IS_NULLABLE"),
                    CharacterMaximumLength: TryGetNullableInt(row, "CHARACTER_MAXIMUM_LENGTH")));
            }

            return (IReadOnlyList<ColumnInfo>)columns
                .Where(c => !string.IsNullOrWhiteSpace(c.Name))
                .OrderBy(c => c.Ordinal)
                .ThenBy(c => c.Name, StringComparer.OrdinalIgnoreCase)
                .ToArray();

            static int TryGetInt(DataRow row, string name)
                => row.Table.Columns.Contains(name) && row[name] is not DBNull ? Convert.ToInt32(row[name]) : 0;

            static int? TryGetNullableInt(DataRow row, string name)
                => row.Table.Columns.Contains(name) && row[name] is not DBNull ? Convert.ToInt32(row[name]) : null;

            static bool TryGetBool(DataRow row, string name)
            {
                if (!row.Table.Columns.Contains(name) || row[name] is DBNull)
                {
                    return false;
                }

                var value = row[name];
                return value switch
                {
                    bool b => b,
                    int i => i != 0,
                    short s => s != 0,
                    string str => str.Equals("true", StringComparison.OrdinalIgnoreCase) || str.Equals("yes", StringComparison.OrdinalIgnoreCase),
                    _ => Convert.ToBoolean(value)
                };
            }
        }), cancellationToken);
    }

    /// <summary>
    /// Executes a query and returns a JSON-friendly row/column shape.
    /// </summary>
    public Task<QueryResult> QueryAsync(string databasePath, string sql, IReadOnlyDictionary<string, object?>? parameters, CancellationToken cancellationToken)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(sql);

        return Task.Run(() => OleDbStaThread.Execute(() =>
        {
            using var connection = CreateConnection(databasePath);
            connection.Open();

            using var command = connection.CreateCommand();
            command.CommandText = sql;
            command.CommandTimeout = _options.CommandTimeoutSeconds;
            AddParameters(command, parameters);

            using var reader = command.ExecuteReader();

            // Read a row-oriented JSON-friendly shape.
            var columns = Enumerable.Range(0, reader.FieldCount)
                .Select(i => reader.GetName(i))
                .ToArray();

            var rows = new List<Dictionary<string, object?>>();
            while (reader.Read())
            {
                var row = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
                for (var i = 0; i < columns.Length; i++)
                {
                    var value = reader.IsDBNull(i) ? null : reader.GetValue(i);
                    row[columns[i]] = value;
                }
                rows.Add(row);
            }

            return new QueryResult((IReadOnlyList<string>)columns, rows);
        }), cancellationToken);
    }

    /// <summary>
    /// Executes a scalar query and returns the single value.
    /// </summary>
    public Task<object?> QueryScalarAsync(string databasePath, string sql, IReadOnlyDictionary<string, object?>? parameters, CancellationToken cancellationToken)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(sql);

        return Task.Run(() => OleDbStaThread.Execute(() =>
        {
            using var connection = CreateConnection(databasePath);
            connection.Open();

            using var command = connection.CreateCommand();
            command.CommandText = sql;
            command.CommandTimeout = _options.CommandTimeoutSeconds;
            AddParameters(command, parameters);

            var value = command.ExecuteScalar();
            return value is DBNull ? null : value;
        }), cancellationToken);
    }

    /// <summary>
    /// Executes a non-query statement and returns the affected row count.
    /// </summary>
    public Task<int> ExecuteAsync(string databasePath, string sql, IReadOnlyDictionary<string, object?>? parameters, CancellationToken cancellationToken)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(sql);

        return Task.Run(() => OleDbStaThread.Execute(() =>
        {
            using var connection = CreateConnection(databasePath);
            connection.Open();

            using var command = connection.CreateCommand();
            command.CommandText = sql;
            command.CommandTimeout = _options.CommandTimeoutSeconds;
            AddParameters(command, parameters);

            return command.ExecuteNonQuery();
        }), cancellationToken);
    }

    private static void AddParameters(OleDbCommand command, IReadOnlyDictionary<string, object?>? parameters)
    {
        if (parameters is null || parameters.Count == 0)
        {
            return;
        }

        // OLE DB uses positional parameters; however, callers commonly pass named params.
        // We'll add in dictionary enumeration order; caller should ensure placeholders match.
        foreach (var (name, value) in parameters)
        {
            var p = command.CreateParameter();
            p.ParameterName = name;
            p.Value = value ?? DBNull.Value;
            command.Parameters.Add(p);
        }
    }
}

/// <summary>
/// Column metadata returned from Access schema queries.
/// </summary>
internal sealed record ColumnInfo(string Name, int Ordinal, string DataType, bool IsNullable, int? CharacterMaximumLength);

/// <summary>
/// Index metadata returned from Access schema queries.
/// </summary>
internal sealed record IndexInfo(string Name, string Table, bool IsPrimaryKey, bool IsUnique, string Type);

/// <summary>
/// Foreign key metadata returned from Access schema queries.
/// </summary>
internal sealed record RelationshipInfo(string Name, string PrimaryTable, string PrimaryColumn, string ForeignTable, string ForeignColumn, string UpdateRule, string DeleteRule);

/// <summary>
/// Query result shape with columns and row data for JSON serialization.
/// </summary>
internal sealed record QueryResult(IReadOnlyList<string> Columns, IReadOnlyList<IReadOnlyDictionary<string, object?>> Rows);
