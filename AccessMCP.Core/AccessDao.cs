using System.Globalization;

namespace DigitalJeeves.MsAccessMCP.Core;

/// <summary>
/// Runs DAO-only operations directly against DAO.DBEngine, without ever launching Access.Application -
/// so no MSACCESS.EXE process is spawned. Unlike <see cref="AccessComAutomation"/>, which automates the
/// full Access application shell (needed for forms/reports/VBE work), this is for metadata that DAO alone
/// can answer.
/// <para>
/// DAO and the ACE OLE DB provider used by <see cref="AccessOleDb"/> load the same underlying engine DLL
/// in-process. This is routed through <see cref="OleDbStaThread"/> - the same dedicated STA thread
/// <see cref="AccessOleDb"/> already uses - rather than a thread of its own, so the engine is never
/// entered from two apartments at once. See <see cref="OleDbStaThread"/>'s doc comment for why that
/// matters for this provider.
/// </para>
/// </summary>
internal sealed class AccessDao(AccessMcpOptions options)
{
    private readonly AccessMcpOptions _options = options;

    /// <summary>
    /// Lists tables with field and index design metadata via DAO, without launching Access.
    /// </summary>
    public Task<IReadOnlyList<object>> GetTableDesignAsync(string databasePath, string? tableName, CancellationToken cancellationToken) =>
        Task.Run(() => OleDbStaThread.Execute(() =>
        {
            dynamic? dbEngine = null;
            dynamic? db = null;
            dynamic? tableDefs = null;

            try
            {
                var type = Type.GetTypeFromProgID(_options.DaoProgId)
                    ?? throw new InvalidOperationException($"DAO engine '{_options.DaoProgId}' is not registered.");

                dbEngine = Activator.CreateInstance(type)
                    ?? throw new InvalidOperationException("Failed to create DAO.DBEngine COM instance.");

                if (!string.IsNullOrWhiteSpace(_options.WorkgroupPath))
                {
                    dbEngine.SystemDB = _options.WorkgroupPath;
                }

                // Shared mode, read-only - this path only ever reads design metadata.
                db = dbEngine.OpenDatabase(databasePath, false, true);
                tableDefs = db.TableDefs;

                var result = new List<object>();
                var count = (int)tableDefs.Count;
                for (var i = 0; i < count; i++)
                {
                    var tableDef = tableDefs.Item(i);
                    try
                    {
                        var name = Convert.ToString(tableDef.Name) ?? string.Empty;
                        if (string.IsNullOrWhiteSpace(name))
                        {
                            continue;
                        }

                        if (!string.IsNullOrWhiteSpace(tableName)
                            && !name.Equals(tableName, StringComparison.OrdinalIgnoreCase))
                        {
                            continue;
                        }

                        if (name.StartsWith("MSys", StringComparison.OrdinalIgnoreCase) || name.StartsWith("~", StringComparison.OrdinalIgnoreCase))
                        {
                            continue;
                        }

                        result.Add(ReadTableDesign(tableDef, name));
                    }
                    finally
                    {
                        AccessComAutomation.ReleaseComObject(tableDef);
                    }
                }

                return (IReadOnlyList<object>)result;
            }
            finally
            {
                AccessComAutomation.ReleaseComObject(tableDefs);

                if (db is not null)
                {
                    try
                    {
                        db.Close();
                    }
                    catch
                    {
                    }
                }

                AccessComAutomation.ReleaseComObject(db);
                AccessComAutomation.ReleaseComObject(dbEngine);
            }
        }), cancellationToken);

    private static object ReadTableDesign(dynamic tableDef, string name)
    {
        var fields = new List<object>();
        var indexes = new List<object>();

        dynamic? fieldsCollection = null;
        dynamic? indexesCollection = null;

        try
        {
            fieldsCollection = tableDef.Fields;
            var fieldCount = (int)fieldsCollection.Count;
            for (var f = 0; f < fieldCount; f++)
            {
                var field = fieldsCollection.Item(f);
                try
                {
                    fields.Add(new
                    {
                        name = Convert.ToString(field.Name) ?? string.Empty,
                        type = Convert.ToInt32(field.Type, CultureInfo.InvariantCulture),
                        size = Convert.ToInt32(field.Size, CultureInfo.InvariantCulture),
                        required = Convert.ToBoolean(field.Required, CultureInfo.InvariantCulture)
                    });
                }
                finally
                {
                    AccessComAutomation.ReleaseComObject(field);
                }
            }

            indexesCollection = tableDef.Indexes;
            var indexCount = (int)indexesCollection.Count;
            for (var ix = 0; ix < indexCount; ix++)
            {
                var index = indexesCollection.Item(ix);
                try
                {
                    var fieldNames = new List<string>();
                    dynamic? indexFields = null;
                    try
                    {
                        indexFields = index.Fields;
                        var indexFieldCount = (int)indexFields.Count;
                        for (var fx = 0; fx < indexFieldCount; fx++)
                        {
                            var indexField = indexFields.Item(fx);
                            try
                            {
                                var fieldName = Convert.ToString(indexField.Name);
                                if (!string.IsNullOrWhiteSpace(fieldName))
                                {
                                    fieldNames.Add(fieldName);
                                }
                            }
                            finally
                            {
                                AccessComAutomation.ReleaseComObject(indexField);
                            }
                        }
                    }
                    finally
                    {
                        AccessComAutomation.ReleaseComObject(indexFields);
                    }

                    indexes.Add(new
                    {
                        name = Convert.ToString(index.Name) ?? string.Empty,
                        primary = Convert.ToBoolean(index.Primary, CultureInfo.InvariantCulture),
                        unique = Convert.ToBoolean(index.Unique, CultureInfo.InvariantCulture),
                        fields = fieldNames
                    });
                }
                finally
                {
                    AccessComAutomation.ReleaseComObject(index);
                }
            }
        }
        finally
        {
            AccessComAutomation.ReleaseComObject(fieldsCollection);
            AccessComAutomation.ReleaseComObject(indexesCollection);
        }

        return new { name, fields, indexes };
    }
}
