using ModelContextProtocol.Server;
using System.ComponentModel;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace DigitalJeeves.MsAccessMCP.Core;

/// <summary>
/// MCP tool surface for Access databases, exposing OLE DB and COM automation operations.
/// </summary>
[McpServerToolType]
public sealed class AccessTools(AccessMcpOptions options)
{
    private readonly AccessMcpOptions _options = options;

    [McpServerTool(Name = "access_list_tables", Title = "Access: List Tables", ReadOnly = true, Idempotent = true)]
    [Description("Lists user tables in an Access database.")]
    public async Task<string> ListTablesAsync(
        [Description("Path to .mdb/.accdb. If omitted, uses DefaultDatabasePath.")] string? databasePath = null,
        CancellationToken cancellationToken = default)
    {
        var dbPath = AccessDbPathPolicy.NormalizeAndValidate(databasePath, _options);

        var db = new AccessOleDb(_options);
        var tables = await db.ListTablesAsync(dbPath, cancellationToken);

        return JsonSerializer.Serialize(new { databasePath = dbPath, tables }, JsonOptions);
    }

    [McpServerTool(Name = "access_list_views", Title = "Access: List Views", ReadOnly = true, Idempotent = true)]
    [Description("Lists views in an Access database.")]
    public async Task<string> ListViewsAsync(
        [Description("Path to .mdb/.accdb. If omitted, uses DefaultDatabasePath.")] string? databasePath = null,
        CancellationToken cancellationToken = default)
    {
        var dbPath = AccessDbPathPolicy.NormalizeAndValidate(databasePath, _options);

        var db = new AccessOleDb(_options);
        var views = await db.ListViewsAsync(dbPath, cancellationToken);

        return JsonSerializer.Serialize(new { databasePath = dbPath, views }, JsonOptions);
    }

    /// <summary>
    /// Lists stored queries in an Access database.
    /// </summary>
    [McpServerTool(Name = "access_list_queries", Title = "Access: List Queries", ReadOnly = true, Idempotent = true)]
    [Description("Lists stored queries in an Access database.")]
    public async Task<string> ListQueriesAsync(
        [Description("Path to .mdb/.accdb. If omitted, uses DefaultDatabasePath.")] string? databasePath = null,
        CancellationToken cancellationToken = default)
    {
        var dbPath = AccessDbPathPolicy.NormalizeAndValidate(databasePath, _options);

        var db = new AccessOleDb(_options);
        var queries = await db.ListQueriesAsync(dbPath, cancellationToken);

        return JsonSerializer.Serialize(new { databasePath = dbPath, queries }, JsonOptions);
    }

    /// <summary>
    /// Lists indexes in an Access database.
    /// </summary>
    [McpServerTool(Name = "access_list_indexes", Title = "Access: List Indexes", ReadOnly = true, Idempotent = true)]
    [Description("Lists indexes in an Access database.")]
    public async Task<string> ListIndexesAsync(
        [Description("Path to .mdb/.accdb. If omitted, uses DefaultDatabasePath.")] string? databasePath = null,
        [Description("Optional table name filter.")] string? tableName = null,
        CancellationToken cancellationToken = default)
    {
        var dbPath = AccessDbPathPolicy.NormalizeAndValidate(databasePath, _options);

        var db = new AccessOleDb(_options);
        var indexes = await db.ListIndexesAsync(dbPath, tableName, cancellationToken);

        return JsonSerializer.Serialize(new { databasePath = dbPath, indexes }, JsonOptions);
    }

    /// <summary>
    /// Lists relationships in an Access database.
    /// </summary>
    [McpServerTool(Name = "access_list_relationships", Title = "Access: List Relationships", ReadOnly = true, Idempotent = true)]
    [Description("Lists relationships in an Access database.")]
    public async Task<string> ListRelationshipsAsync(
        [Description("Path to .mdb/.accdb. If omitted, uses DefaultDatabasePath.")] string? databasePath = null,
        [Description("Optional table name filter.")] string? tableName = null,
        CancellationToken cancellationToken = default)
    {
        var dbPath = AccessDbPathPolicy.NormalizeAndValidate(databasePath, _options);

        var db = new AccessOleDb(_options);
        var relationships = await db.ListRelationshipsAsync(dbPath, tableName, cancellationToken);

        return JsonSerializer.Serialize(new { databasePath = dbPath, relationships }, JsonOptions);
    }

    [McpServerTool(Name = "access_get_schema", Title = "Access: Get Schema", ReadOnly = true, Idempotent = true)]
    [Description("Returns tables and columns for an Access database.")]
    public async Task<string> GetSchemaAsync(
        [Description("Path to .mdb/.accdb. If omitted, uses DefaultDatabasePath.")] string? databasePath = null,
        [Description("Optional table name filter.")] string? tableName = null,
        CancellationToken cancellationToken = default)
    {
        var dbPath = AccessDbPathPolicy.NormalizeAndValidate(databasePath, _options);

        var db = new AccessOleDb(_options);
        var tables = string.IsNullOrWhiteSpace(tableName)
            ? await db.ListTablesAsync(dbPath, cancellationToken)
            : [tableName];

        var result = new List<object>(tables.Count);
        foreach (var t in tables)
        {
            var cols = await db.ListColumnsAsync(dbPath, t, cancellationToken);
            result.Add(new { table = t, columns = cols });
        }

        return JsonSerializer.Serialize(new { databasePath = dbPath, schema = result }, JsonOptions);
    }

    [McpServerTool(Name = "access_query", Title = "Access: Query", ReadOnly = true, Idempotent = true)]
    [Description("Executes a SELECT query against an Access database and returns rows.")]
    public async Task<string> QueryAsync(
        [Description("SQL query text (typically SELECT).")] string sql,
        [Description("Path to .mdb/.accdb. If omitted, uses DefaultDatabasePath.")] string? databasePath = null,
        [Description("Optional parameters as a JSON object. Note: OLE DB uses positional parameters.")] JsonElement? parameters = null,
        CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(sql);

        var dbPath = AccessDbPathPolicy.NormalizeAndValidate(databasePath, _options);
        var args = ParseParameters(parameters);

        var db = new AccessOleDb(_options);
        var result = await db.QueryAsync(dbPath, sql, args, cancellationToken);

        return JsonSerializer.Serialize(new { databasePath = dbPath, result }, JsonOptions);
    }

    /// <summary>
    /// Executes a SELECT query and returns a paged slice of the result set.
    /// </summary>
    [McpServerTool(Name = "access_query_paged", Title = "Access: Query (Paged)", ReadOnly = true, Idempotent = true)]
    [Description("Executes a SELECT query and returns a paged slice (client-side paging).")]
    public async Task<string> QueryPagedAsync(
        [Description("SQL query text (typically SELECT).")] string sql,
        [Description("1-based page number.")] int pageNumber,
        [Description("Page size (rows per page).")] int pageSize,
        [Description("Path to .mdb/.accdb. If omitted, uses DefaultDatabasePath.")] string? databasePath = null,
        [Description("Optional parameters as a JSON object. Note: OLE DB uses positional parameters.")] JsonElement? parameters = null,
        CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(sql);

        if (pageNumber < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(pageNumber), "Page number must be at least 1.");
        }

        if (pageSize <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(pageSize), "Page size must be greater than zero.");
        }

        var dbPath = AccessDbPathPolicy.NormalizeAndValidate(databasePath, _options);
        var args = ParseParameters(parameters);

        var db = new AccessOleDb(_options);
        var result = await db.QueryAsync(dbPath, sql, args, cancellationToken);

        var totalRows = result.Rows.Count;
        var skip = (pageNumber - 1) * pageSize;
        var pageRows = new List<IReadOnlyDictionary<string, object?>>(Math.Min(pageSize, Math.Max(0, totalRows - skip)));

        for (var i = skip; i < totalRows && pageRows.Count < pageSize; i++)
        {
            pageRows.Add(result.Rows[i]);
        }

        var payload = new
        {
            databasePath = dbPath,
            pageNumber,
            pageSize,
            totalRows,
            columns = result.Columns,
            rows = pageRows
        };

        return JsonSerializer.Serialize(payload, JsonOptions);
    }

    /// <summary>
    /// Executes a scalar query against an Access database and returns a single value.
    /// </summary>
    [McpServerTool(Name = "access_scalar", Title = "Access: Scalar", ReadOnly = true, Idempotent = true)]
    [Description("Executes a scalar query (single value) against an Access database.")]
    public async Task<string> ScalarAsync(
        [Description("SQL query text (typically SELECT with aggregate).")] string sql,
        [Description("Path to .mdb/.accdb. If omitted, uses DefaultDatabasePath.")] string? databasePath = null,
        [Description("Optional parameters as a JSON object. Note: OLE DB uses positional parameters.")] JsonElement? parameters = null,
        CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(sql);

        var dbPath = AccessDbPathPolicy.NormalizeAndValidate(databasePath, _options);
        var args = ParseParameters(parameters);

        var db = new AccessOleDb(_options);
        var value = await db.QueryScalarAsync(dbPath, sql, args, cancellationToken);

        return JsonSerializer.Serialize(new { databasePath = dbPath, value }, JsonOptions);
    }

    /// <summary>
    /// Exports a table as JSON or CSV.
    /// </summary>
    [McpServerTool(Name = "access_export_table", Title = "Access: Export Table", ReadOnly = true, Idempotent = true)]
    [Description("Exports a table as JSON or CSV.")]
    public async Task<string> ExportTableAsync(
        [Description("Table name to export.")] string tableName,
        [Description("Export format: json or csv.")] string format,
        [Description("Optional max rows to export.")] int? maxRows = null,
        [Description("Path to .mdb/.accdb. If omitted, uses DefaultDatabasePath.")] string? databasePath = null,
        CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(tableName);
        ArgumentException.ThrowIfNullOrWhiteSpace(format);

        if (maxRows is not null && maxRows <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(maxRows), "Max rows must be greater than zero.");
        }

        var normalizedFormat = format.Trim();
        if (!normalizedFormat.Equals("json", StringComparison.OrdinalIgnoreCase)
            && !normalizedFormat.Equals("csv", StringComparison.OrdinalIgnoreCase))
        {
            throw new ArgumentException("Format must be 'json' or 'csv'.", nameof(format));
        }

        var dbPath = AccessDbPathPolicy.NormalizeAndValidate(databasePath, _options);
        var sql = $"SELECT * FROM [{tableName}]";

        var db = new AccessOleDb(_options);
        var result = await db.QueryAsync(dbPath, sql, null, cancellationToken);

        var rows = maxRows is null
            ? result.Rows
            : result.Rows.Take(maxRows.Value).ToArray();

        if (normalizedFormat.Equals("csv", StringComparison.OrdinalIgnoreCase))
        {
            var csv = SerializeCsv(result.Columns, rows);
            return JsonSerializer.Serialize(new { databasePath = dbPath, tableName, format = "csv", csv }, JsonOptions);
        }

        return JsonSerializer.Serialize(new { databasePath = dbPath, tableName, format = "json", columns = result.Columns, rows }, JsonOptions);
    }

    [McpServerTool(Name = "access_execute", Title = "Access: Execute", ReadOnly = false, Idempotent = false)]
    [Description("Executes a non-query statement (INSERT/UPDATE/DELETE/DDL).")]
    public async Task<string> ExecuteAsync(
        [Description("SQL statement text.")] string sql,
        [Description("Path to .mdb/.accdb. If omitted, uses DefaultDatabasePath.")] string? databasePath = null,
        [Description("Optional parameters as a JSON object. Note: OLE DB uses positional parameters.")] JsonElement? parameters = null,
        CancellationToken cancellationToken = default)
    {
        if (_options.ReadOnly)
        {
            throw new InvalidOperationException("Server is configured as read-only.");
        }

        ArgumentException.ThrowIfNullOrWhiteSpace(sql);

        var dbPath = AccessDbPathPolicy.NormalizeAndValidate(databasePath, _options);
        var args = ParseParameters(parameters);

        var db = new AccessOleDb(_options);
        var affected = await db.ExecuteAsync(dbPath, sql, args, cancellationToken);

        return JsonSerializer.Serialize(new { databasePath = dbPath, affected }, JsonOptions);
    }

    /// <summary>
    /// Lists Access forms via COM automation.
    /// </summary>
    [McpServerTool(Name = "access_list_forms", Title = "Access: List Forms", ReadOnly = true, Idempotent = true)]
    [Description("Lists Access forms via COM automation.")]
    public async Task<string> ListFormsAsync(
        [Description("Path to .mdb/.accdb. If omitted, uses DefaultDatabasePath.")] string? databasePath = null,
        CancellationToken cancellationToken = default)
    {
        var dbPath = AccessDbPathPolicy.NormalizeAndValidate(databasePath, _options);
        var com = new AccessComAutomation(_options);

        var forms = await com.RunInAccessAsync(dbPath, app =>
        {
            var formNames = new List<string>();
            var allForms = app.CurrentProject.AllForms;

            try
            {
                var count = (int)allForms.Count;
                for (var i = 0; i < count; i++)
                {
                    var formRef = allForms.Item(i);
                    try
                    {
                        var name = Convert.ToString(formRef.Name);
                        if (!string.IsNullOrWhiteSpace(name))
                        {
                            formNames.Add(name);
                        }
                    }
                    finally
                    {
                        AccessComAutomation.ReleaseComObject(formRef);
                    }
                }
            }
            finally
            {
                AccessComAutomation.ReleaseComObject(allForms);
            }

            formNames.Sort(StringComparer.OrdinalIgnoreCase);
            return (IReadOnlyList<string>)formNames;
        }, cancellationToken);

        return JsonSerializer.Serialize(new { databasePath = dbPath, forms }, JsonOptions);
    }

    /// <summary>
    /// Gets controls for an Access form via COM automation.
    /// </summary>
    [McpServerTool(Name = "access_get_form_controls", Title = "Access: Get Form Controls", ReadOnly = true, Idempotent = true)]
    [Description("Gets controls for an Access form via COM automation.")]
    public async Task<string> GetFormControlsAsync(
        [Description("Form name.")] string formName,
        [Description("If true, includes current control values (form view).")]
        bool includeValues = false,
        [Description("Path to .mdb/.accdb. If omitted, uses DefaultDatabasePath.")] string? databasePath = null,
        CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(formName);

        var dbPath = AccessDbPathPolicy.NormalizeAndValidate(databasePath, _options);
        var com = new AccessComAutomation(_options);
        var view = AccessComAutomation.ResolveFormView(!includeValues);
        var closeSave = AccessComAutomation.ResolveCloseSave(false);

        var controls = await com.RunInAccessAsync(dbPath, app =>
        {
            var controlInfos = new List<object>();
            var doCmd = app.DoCmd;
            doCmd.OpenForm(formName, view);

            // Held as `object`, not `dynamic` — see the note on AccessComAutomation's reflection helpers.
            object? form = null;
            object? controlsCollection = null;

            try
            {
                form = app.Forms[formName];
                controlsCollection = AccessComAutomation.GetProperty(form!, "Controls")!;

                var count = AccessComAutomation.GetCount(controlsCollection);
                for (var i = 0; i < count; i++)
                {
                    var control = AccessComAutomation.GetItem(controlsCollection, i)!;
                    try
                    {
                        var name = Convert.ToString(AccessComAutomation.GetProperty(control, "Name")) ?? string.Empty;
                        var controlType = Convert.ToInt32(AccessComAutomation.GetProperty(control, "ControlType"), CultureInfo.InvariantCulture);
                        var controlSource = Convert.ToString(AccessComAutomation.TryGetProperty(control, "ControlSource"));
                        var value = includeValues ? AccessComAutomation.TryGetProperty(control, "Value") : null;

                        controlInfos.Add(new
                        {
                            name,
                            controlType,
                            controlTypeName = AccessComAutomation.GetControlTypeName(controlType),
                            controlSource,
                            value
                        });
                    }
                    finally
                    {
                        AccessComAutomation.ReleaseComObject(control);
                    }
                }
            }
            finally
            {
                try
                {
                    doCmd.Close(AccessComAutomation.FormObjectType, formName, closeSave);
                }
                catch
                {
                }

                AccessComAutomation.ReleaseComObject(controlsCollection);
                AccessComAutomation.ReleaseComObject(form);
                AccessComAutomation.ReleaseComObject(doCmd);
            }

            return (IReadOnlyList<object>)controlInfos;
        }, cancellationToken);

        return JsonSerializer.Serialize(new { databasePath = dbPath, formName, controls }, JsonOptions);
    }

    /// <summary>
    /// Updates a control value on an Access form via COM automation.
    /// </summary>
    [McpServerTool(Name = "access_set_form_control_value", Title = "Access: Set Form Control Value", ReadOnly = false, Idempotent = false)]
    [Description("Updates a control value on an Access form via COM automation.")]
    public async Task<string> SetFormControlValueAsync(
        [Description("Form name.")] string formName,
        [Description("Control name.")] string controlName,
        [Description("Value to assign (JSON value).")]
        JsonElement? value = null,
        [Description("If true, saves the current record.")]
        bool saveRecord = true,
        [Description("Path to .mdb/.accdb. If omitted, uses DefaultDatabasePath.")] string? databasePath = null,
        CancellationToken cancellationToken = default)
    {
        if (_options.ReadOnly)
        {
            throw new InvalidOperationException("Server is configured as read-only.");
        }

        ArgumentException.ThrowIfNullOrWhiteSpace(formName);
        ArgumentException.ThrowIfNullOrWhiteSpace(controlName);

        var dbPath = AccessDbPathPolicy.NormalizeAndValidate(databasePath, _options);
        var com = new AccessComAutomation(_options);
        var closeSave = AccessComAutomation.ResolveCloseSave(false);
        var typedValue = ParseJsonValue(value);

        await com.RunInAccessAsync(dbPath, app =>
        {
            var doCmd = app.DoCmd;
            doCmd.OpenForm(formName, AccessComAutomation.ResolveFormView(false));

            // Held as `object`, not `dynamic` — see the note on AccessComAutomation's reflection helpers.
            object? form = null;
            object? control = null;

            try
            {
                form = app.Forms[formName];
                control = AccessComAutomation.GetItem(AccessComAutomation.GetProperty(form!, "Controls")!, controlName)!;
                AccessComAutomation.SetProperty(control, "Value", typedValue ?? DBNull.Value);

                if (saveRecord)
                {
                    doCmd.RunCommand(AccessComAutomation.SaveRecordCommand);
                }
            }
            finally
            {
                try
                {
                    doCmd.Close(AccessComAutomation.FormObjectType, formName, closeSave);
                }
                catch
                {
                }

                AccessComAutomation.ReleaseComObject(control);
                AccessComAutomation.ReleaseComObject(form);
                AccessComAutomation.ReleaseComObject(doCmd);
            }

            return true;
        }, cancellationToken);

        return JsonSerializer.Serialize(new { databasePath = dbPath, formName, controlName, value = typedValue, saveRecord }, JsonOptions);
    }

    /// <summary>
    /// Adds a control to an Access form via COM automation.
    /// </summary>
    [McpServerTool(Name = "access_add_form_control", Title = "Access: Add Form Control", ReadOnly = false, Idempotent = false)]
    [Description("Adds a control to an Access form via COM automation.")]
    public async Task<string> AddFormControlAsync(
        [Description("Form name.")] string formName,
        [Description("Control type (e.g., textbox, combobox, checkbox).")]
        string controlType,
        [Description("Control name to assign.")] string controlName,
        [Description("Section (detail, header, footer).")]
        string section = "detail",
        [Description("Left position (twips).")]
        double? left = null,
        [Description("Top position (twips).")]
        double? top = null,
        [Description("Width (twips).")]
        double? width = null,
        [Description("Height (twips).")]
        double? height = null,
        [Description("Path to .mdb/.accdb. If omitted, uses DefaultDatabasePath.")] string? databasePath = null,
        CancellationToken cancellationToken = default)
    {
        if (_options.ReadOnly)
        {
            throw new InvalidOperationException("Server is configured as read-only.");
        }

        ArgumentException.ThrowIfNullOrWhiteSpace(formName);
        ArgumentException.ThrowIfNullOrWhiteSpace(controlType);
        ArgumentException.ThrowIfNullOrWhiteSpace(controlName);

        var dbPath = AccessDbPathPolicy.NormalizeAndValidate(databasePath, _options);
        var com = new AccessComAutomation(_options);
        var controlTypeValue = AccessComAutomation.ResolveControlType(controlType);
        var sectionValue = AccessComAutomation.ResolveSection(section);
        var closeSave = AccessComAutomation.ResolveCloseSave(true);

        await com.RunInAccessAsync(dbPath, app =>
        {
            var doCmd = app.DoCmd;
            doCmd.OpenForm(formName, AccessComAutomation.ResolveFormView(true));

            dynamic? created = null;

            try
            {
                created = app.CreateControl(
                    formName,
                    controlTypeValue,
                    sectionValue,
                    Type.Missing,
                    Type.Missing,
                    left ?? Type.Missing,
                    top ?? Type.Missing,
                    width ?? Type.Missing,
                    height ?? Type.Missing);

                created.Name = controlName;
            }
            finally
            {
                try
                {
                    doCmd.Close(AccessComAutomation.FormObjectType, formName, closeSave);
                }
                catch
                {
                }

                AccessComAutomation.ReleaseComObject(created);
                AccessComAutomation.ReleaseComObject(doCmd);
            }

            return true;
        }, cancellationToken);

        return JsonSerializer.Serialize(new { databasePath = dbPath, formName, controlName, controlType = controlTypeValue, section = sectionValue }, JsonOptions);
    }

    /// <summary>
    /// Removes a control from an Access form via COM automation.
    /// </summary>
    [McpServerTool(Name = "access_remove_form_control", Title = "Access: Remove Form Control", ReadOnly = false, Idempotent = false)]
    [Description("Removes a control from an Access form via COM automation.")]
    public async Task<string> RemoveFormControlAsync(
        [Description("Form name.")] string formName,
        [Description("Control name.")] string controlName,
        [Description("Path to .mdb/.accdb. If omitted, uses DefaultDatabasePath.")] string? databasePath = null,
        CancellationToken cancellationToken = default)
    {
        if (_options.ReadOnly)
        {
            throw new InvalidOperationException("Server is configured as read-only.");
        }

        ArgumentException.ThrowIfNullOrWhiteSpace(formName);
        ArgumentException.ThrowIfNullOrWhiteSpace(controlName);

        var dbPath = AccessDbPathPolicy.NormalizeAndValidate(databasePath, _options);
        var com = new AccessComAutomation(_options);
        var closeSave = AccessComAutomation.ResolveCloseSave(true);

        await com.RunInAccessAsync(dbPath, app =>
        {
            var doCmd = app.DoCmd;
            doCmd.OpenForm(formName, AccessComAutomation.ResolveFormView(true));

            try
            {
                app.DeleteControl(formName, controlName);
            }
            finally
            {
                try
                {
                    doCmd.Close(AccessComAutomation.FormObjectType, formName, closeSave);
                }
                catch
                {
                }

                AccessComAutomation.ReleaseComObject(doCmd);
            }

            return true;
        }, cancellationToken);

        return JsonSerializer.Serialize(new { databasePath = dbPath, formName, controlName }, JsonOptions);
    }

    [McpServerTool(Name = "access_list_modules", Title = "Access: List Modules", ReadOnly = true, Idempotent = true)]
    [Description("Lists VBA modules in an Access database via COM automation.")]
    public async Task<string> ListModulesAsync(
        [Description("Path to .mdb/.accdb. If omitted, uses DefaultDatabasePath.")] string? databasePath = null,
        CancellationToken cancellationToken = default)
    {
        var dbPath = AccessDbPathPolicy.NormalizeAndValidate(databasePath, _options);
        var com = new AccessComAutomation(_options);

        var modules = await com.RunInAccessAsync(dbPath, app =>
        {
            var names = new List<string>();
            dynamic? allModules = null;

            try
            {
                allModules = app.CurrentProject.AllModules;
                var count = (int)allModules.Count;
                for (var i = 0; i < count; i++)
                {
                    var moduleRef = allModules.Item(i);
                    try
                    {
                        var name = Convert.ToString(moduleRef.Name);
                        if (!string.IsNullOrWhiteSpace(name))
                        {
                            names.Add(name);
                        }
                    }
                    finally
                    {
                        AccessComAutomation.ReleaseComObject(moduleRef);
                    }
                }
            }
            finally
            {
                AccessComAutomation.ReleaseComObject(allModules);
            }

            names.Sort(StringComparer.OrdinalIgnoreCase);
            return (IReadOnlyList<string>)names;
        }, cancellationToken);

        return JsonSerializer.Serialize(new { databasePath = dbPath, modules }, JsonOptions);
    }

    /// <summary>
    /// Lists VBA projects via the VBE object model.
    /// </summary>
    [McpServerTool(Name = "access_vbe_list_projects", Title = "Access: VBE List Projects", ReadOnly = true, Idempotent = true)]
    [Description("Lists VBA projects via the VBE object model.")]
    public async Task<string> ListVbeProjectsAsync(
        [Description("Path to .mdb/.accdb. If omitted, uses DefaultDatabasePath.")] string? databasePath = null,
        CancellationToken cancellationToken = default)
    {
        var dbPath = AccessDbPathPolicy.NormalizeAndValidate(databasePath, _options);
        var com = new AccessComAutomation(_options);

        var projects = await com.RunInAccessAsync(dbPath, app =>
        {
            dynamic? vbe = null;
            dynamic? projectsCollection = null;

            try
            {
                vbe = app.VBE;
                projectsCollection = vbe.VBProjects;

                var names = new List<string>();
                var count = (int)projectsCollection.Count;
                // VBIDE collections are 1-based, unlike Access's own AllForms/AllModules/Controls.
                for (var i = 1; i <= count; i++)
                {
                    var project = projectsCollection.Item(i);
                    try
                    {
                        var name = Convert.ToString(project.Name);
                        if (!string.IsNullOrWhiteSpace(name))
                        {
                            names.Add(name);
                        }
                    }
                    finally
                    {
                        AccessComAutomation.ReleaseComObject(project);
                    }
                }

                names.Sort(StringComparer.OrdinalIgnoreCase);
                return (IReadOnlyList<string>)names;
            }
            finally
            {
                AccessComAutomation.ReleaseComObject(projectsCollection);
                AccessComAutomation.ReleaseComObject(vbe);
            }
        }, cancellationToken);

        return JsonSerializer.Serialize(new { databasePath = dbPath, projects }, JsonOptions);
    }

    /// <summary>
    /// Lists VBA modules via the VBE object model.
    /// </summary>
    [McpServerTool(Name = "access_vbe_list_modules", Title = "Access: VBE List Modules", ReadOnly = true, Idempotent = true)]
    [Description("Lists VBA modules via the VBE object model.")]
    public async Task<string> ListVbeModulesAsync(
        [Description("Optional VBA project name. If omitted, uses the first project.")] string? projectName = null,
        [Description("Path to .mdb/.accdb. If omitted, uses DefaultDatabasePath.")] string? databasePath = null,
        CancellationToken cancellationToken = default)
    {
        var dbPath = AccessDbPathPolicy.NormalizeAndValidate(databasePath, _options);
        var com = new AccessComAutomation(_options);

        var result = await com.RunInAccessAsync(dbPath, app =>
        {
            dynamic? vbe = null;
            dynamic? project = null;
            dynamic? components = null;

            try
            {
                vbe = app.VBE;
                project = GetVbeProject(vbe, projectName);
                components = project.VBComponents;

                var items = new List<object>();
                var count = (int)components.Count;
                // VBIDE collections are 1-based, unlike Access's own AllForms/AllModules/Controls.
                for (var i = 1; i <= count; i++)
                {
                    var component = components.Item(i);
                    dynamic? codeModule = null;
                    try
                    {
                        var name = Convert.ToString(component.Name) ?? string.Empty;
                        int type = Convert.ToInt32(component.Type, CultureInfo.InvariantCulture);
                        codeModule = component.CodeModule;
                        var lineCount = Convert.ToInt32(codeModule.CountOfLines, CultureInfo.InvariantCulture);

                        items.Add(new
                        {
                            name,
                            type,
                            typeName = GetVbeComponentTypeName(type),
                            lineCount
                        });
                    }
                    finally
                    {
                        AccessComAutomation.ReleaseComObject(codeModule);
                        AccessComAutomation.ReleaseComObject(component);
                    }
                }

                var resolvedProjectName = Convert.ToString(project.Name) ?? string.Empty;
                return new { projectName = resolvedProjectName, modules = (IReadOnlyList<object>)items };
            }
            finally
            {
                AccessComAutomation.ReleaseComObject(components);
                AccessComAutomation.ReleaseComObject(project);
                AccessComAutomation.ReleaseComObject(vbe);
            }
        }, cancellationToken);

        return JsonSerializer.Serialize(new { databasePath = dbPath, result.projectName, result.modules }, JsonOptions);
    }

    /// <summary>
    /// Gets full module code via the VBE object model.
    /// </summary>
    [McpServerTool(Name = "access_vbe_get_module_code", Title = "Access: VBE Get Module Code", ReadOnly = true, Idempotent = true)]
    [Description("Gets module code via the VBE object model.")]
    public async Task<string> GetVbeModuleCodeAsync(
        [Description("Module name.")] string moduleName,
        [Description("Optional VBA project name. If omitted, uses the first project.")] string? projectName = null,
        [Description("Path to .mdb/.accdb. If omitted, uses DefaultDatabasePath.")] string? databasePath = null,
        CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(moduleName);

        var dbPath = AccessDbPathPolicy.NormalizeAndValidate(databasePath, _options);
        var com = new AccessComAutomation(_options);

        var result = await com.RunInAccessAsync(dbPath, app =>
        {
            dynamic? vbe = null;
            dynamic? project = null;
            dynamic? component = null;
            dynamic? codeModule = null;

            try
            {
                vbe = app.VBE;
                project = GetVbeProject(vbe, projectName);
                component = project.VBComponents.Item(moduleName);
                codeModule = component.CodeModule;
                var countOfLines = (int)codeModule.CountOfLines;
                var code = countOfLines > 0
                    ? Convert.ToString(codeModule.Lines(1, countOfLines)) ?? string.Empty
                    : string.Empty;

                var resolvedProjectName = Convert.ToString(project.Name) ?? string.Empty;
                return new { projectName = resolvedProjectName, code };
            }
            finally
            {
                AccessComAutomation.ReleaseComObject(codeModule);
                AccessComAutomation.ReleaseComObject(component);
                AccessComAutomation.ReleaseComObject(project);
                AccessComAutomation.ReleaseComObject(vbe);
            }
        }, cancellationToken);

        return JsonSerializer.Serialize(new { databasePath = dbPath, moduleName, result.projectName, result.code }, JsonOptions);
    }

    /// <summary>
    /// Lists procedures in a VBA module via the VBE object model.
    /// </summary>
    [McpServerTool(Name = "access_vbe_get_proc_info", Title = "Access: VBE Get Procedure Info", ReadOnly = true, Idempotent = true)]
    [Description("Lists procedures in a VBA module via the VBE object model.")]
    public async Task<string> GetVbeProcedureInfoAsync(
        [Description("Module name.")] string moduleName,
        [Description("Optional VBA project name. If omitted, uses the first project.")] string? projectName = null,
        [Description("Path to .mdb/.accdb. If omitted, uses DefaultDatabasePath.")] string? databasePath = null,
        CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(moduleName);

        var dbPath = AccessDbPathPolicy.NormalizeAndValidate(databasePath, _options);
        var com = new AccessComAutomation(_options);

        var result = await com.RunInAccessAsync(dbPath, app =>
        {
            dynamic? vbe = null;
            dynamic? project = null;
            dynamic? component = null;
            dynamic? codeModule = null;

            try
            {
                vbe = app.VBE;
                project = GetVbeProject(vbe, projectName);
                component = project.VBComponents.Item(moduleName);
                codeModule = component.CodeModule;
                var countOfLines = (int)codeModule.CountOfLines;

                var items = new List<object>();
                var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                var line = 1;

                while (line <= countOfLines)
                {
                    int kind;
                    var procName = Convert.ToString(codeModule.ProcOfLine(line, out kind));
                    if (string.IsNullOrWhiteSpace(procName))
                    {
                        line++;
                        continue;
                    }

                    var key = $"{procName}:{kind}";
                    if (seen.Add(key))
                    {
                        var startLine = (int)codeModule.ProcStartLine(procName, kind);
                        var lineCount = (int)codeModule.ProcCountLines(procName, kind);
                        items.Add(new
                        {
                            name = procName,
                            kind,
                            startLine,
                            lineCount
                        });
                        line = startLine + Math.Max(1, lineCount);
                    }
                    else
                    {
                        line++;
                    }
                }

                var resolvedProjectName = Convert.ToString(project.Name) ?? string.Empty;
                return new { projectName = resolvedProjectName, procedures = (IReadOnlyList<object>)items };
            }
            finally
            {
                AccessComAutomation.ReleaseComObject(codeModule);
                AccessComAutomation.ReleaseComObject(component);
                AccessComAutomation.ReleaseComObject(project);
                AccessComAutomation.ReleaseComObject(vbe);
            }
        }, cancellationToken);

        return JsonSerializer.Serialize(new { databasePath = dbPath, moduleName, result.projectName, procedures = result.procedures }, JsonOptions);
    }

    /// <summary>
    /// Compares the VBA module code of two Access databases via the VBE object model.
    /// </summary>
    [McpServerTool(Name = "access_diff_vba_modules", Title = "Access: Diff VBA Modules", ReadOnly = true, Idempotent = true)]
    [Description("Compares the VBA module code (and, optionally, the saved queries) of two Access databases via COM automation: which exist only in one, which changed (with a unified diff per item), and which are identical. Both database paths are required. Requires EnableComAutomation and Trust Center \"Trust access to the VBA project object model\".")]
    public async Task<string> DiffVbaModulesAsync(
        [Description("Path to the baseline .mdb/.accdb.")] string sourceDatabasePath,
        [Description("Path to the .mdb/.accdb to compare against the baseline.")] string targetDatabasePath,
        [Description("Optional VBA project name. If omitted, uses the first project in each database.")] string? projectName = null,
        [Description("Optional module name to restrict the comparison to a single module.")] string? moduleName = null,
        [Description("If true, identical modules are listed by name; if false (default) they are only counted in the summary.")] bool includeUnchanged = false,
        [Description("Lines of context around each unified-diff hunk. Default 3.")] int contextLines = 3,
        [Description("If true (the default), differences that are only in letter case (for example 'Web' vs 'web') are ignored: modules whose code differs only by case count as unchanged, and such lines are not shown in the diffs. Set false for exact case-sensitive comparison.")] bool ignoreCase = true,
        [Description("If true (the default), differences that are only in whitespace are ignored: each line is compared after trimming its ends and collapsing internal runs of whitespace to a single space (like git --ignore-space-change). Modules whose code differs only that way count as unchanged, and such lines are not shown in the diffs. Spacing added or removed where there was none (for example 'x=1' vs 'x = 1') is still reported. Set false for exact whitespace comparison. Line-ending style is always ignored. The same rules apply to query SQL.")] bool ignoreWhitespace = true,
        [Description("If true, also compares saved queries (CurrentDb().QueryDefs, including Access's ~ temporary queries and MSys* system queries). Default false: the queries section is omitted and only VBA modules are compared.")] bool includeQueries = false,
        [Description("Optional query name to restrict the query comparison to a single query. Only used when includeQueries is true.")] string? queryName = null,
        CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(sourceDatabasePath);
        ArgumentException.ThrowIfNullOrWhiteSpace(targetDatabasePath);
        if (contextLines < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(contextLines), "Context lines cannot be negative.");
        }

        var sourcePath = AccessDbPathPolicy.NormalizeAndValidate(sourceDatabasePath, _options);
        var targetPath = AccessDbPathPolicy.NormalizeAndValidate(targetDatabasePath, _options);

        // Read sequentially, never concurrently: Access COM automation is effectively single-instance
        // (see AccessComAutomation.AutomationGate), so each database is opened and closed before the next.
        var source = await ReadSnapshotsAsync(sourcePath, projectName, moduleName, includeQueries, queryName, cancellationToken);
        var target = await ReadSnapshotsAsync(targetPath, projectName, moduleName, includeQueries, queryName, cancellationToken);

        var result = VbaModuleDiff.Compare(
            source.Modules,
            target.Modules,
            contextLines,
            ignoreCase,
            ignoreWhitespace,
            sourceLabel: "source",
            targetLabel: "target");

        var queryResult = includeQueries
            ? QueryDiff.Compare(
                source.Queries,
                target.Queries,
                contextLines,
                ignoreCase,
                ignoreWhitespace,
                sourceLabel: "source/queries",
                targetLabel: "target/queries")
            : null;

        object? queries = queryResult is null
            ? null
            : new
            {
                summary = new
                {
                    unchanged = queryResult.Unchanged.Count,
                    changed = queryResult.Changed.Count,
                    onlyInSource = queryResult.OnlyInSource.Count,
                    onlyInTarget = queryResult.OnlyInTarget.Count
                },
                onlyInSource = queryResult.OnlyInSource.Select(ToQuerySummary),
                onlyInTarget = queryResult.OnlyInTarget.Select(ToQuerySummary),
                changed = queryResult.Changed.Select(change => new
                {
                    name = change.Name,
                    type = change.QueryType,
                    typeName = DescribeQueryType(change.QueryType),
                    sourceLines = change.SourceLineCount,
                    targetLines = change.TargetLineCount,
                    diff = change.Diff
                }),
                unchanged = includeUnchanged
                    ? queryResult.Unchanged.Select(query => new
                    {
                        name = query.Name,
                        type = query.QueryType,
                        typeName = DescribeQueryType(query.QueryType)
                    })
                    : null
            };

        var payload = new
        {
            sourceDatabasePath = sourcePath,
            targetDatabasePath = targetPath,
            sourceProjectName = source.ProjectName,
            targetProjectName = target.ProjectName,
            ignoreCase,
            ignoreWhitespace,
            includeQueries,
            summary = new
            {
                unchanged = result.Unchanged.Count,
                changed = result.Changed.Count,
                onlyInSource = result.OnlyInSource.Count,
                onlyInTarget = result.OnlyInTarget.Count
            },
            onlyInSource = result.OnlyInSource.Select(ToModuleSummary),
            onlyInTarget = result.OnlyInTarget.Select(ToModuleSummary),
            changed = result.Changed.Select(change => new
            {
                name = change.Name,
                type = change.ComponentType,
                typeName = GetVbeComponentTypeName(change.ComponentType),
                sourceLines = change.SourceLineCount,
                targetLines = change.TargetLineCount,
                diff = change.Diff
            }),
            unchanged = includeUnchanged
                ? result.Unchanged.Select(module => new
                {
                    name = module.Name,
                    type = module.ComponentType,
                    typeName = GetVbeComponentTypeName(module.ComponentType)
                })
                : null,
            queries
        };

        return JsonSerializer.Serialize(payload, OmitNullsJsonOptions);
    }

    private static object ToQuerySummary(QueryOnly query) => new
    {
        name = query.Name,
        type = query.QueryType,
        typeName = DescribeQueryType(query.QueryType),
        lineCount = query.LineCount
    };

    private static object ToModuleSummary(VbaModuleOnly module) => new
    {
        name = module.Name,
        type = module.ComponentType,
        typeName = GetVbeComponentTypeName(module.ComponentType),
        lineCount = module.LineCount
    };

    /// <summary>
    /// Reads every VBA component (standard, class, and form/report code-behind) of one database, or just
    /// the requested module, via the VBE object model - plus, when requested, every saved query via DAO.
    /// One Access instance serves both reads.
    /// </summary>
    private async Task<DbSnapshotSet> ReadSnapshotsAsync(
        string databasePath,
        string? projectName,
        string? moduleName,
        bool includeQueries,
        string? queryName,
        CancellationToken cancellationToken)
    {
        var com = new AccessComAutomation(_options);

        return await com.RunInAccessAsync(databasePath, app =>
        {
            dynamic? vbe = null;
            dynamic? project = null;
            dynamic? components = null;

            try
            {
                vbe = app.VBE;
                project = GetVbeProject(vbe, projectName);
                components = project.VBComponents;

                var modules = new List<VbaModuleSnapshot>();
                var count = (int)components.Count;
                // VBIDE collections are 1-based, unlike Access's own AllForms/AllModules/Controls.
                for (var i = 1; i <= count; i++)
                {
                    var component = components.Item(i);
                    dynamic? codeModule = null;
                    try
                    {
                        var name = Convert.ToString(component.Name) ?? string.Empty;
                        if (string.IsNullOrWhiteSpace(name))
                        {
                            continue;
                        }

                        if (!string.IsNullOrWhiteSpace(moduleName)
                            && !name.Equals(moduleName, StringComparison.OrdinalIgnoreCase))
                        {
                            continue;
                        }

                        var type = Convert.ToInt32(component.Type, CultureInfo.InvariantCulture);
                        codeModule = component.CodeModule;
                        modules.Add(new VbaModuleSnapshot(name, type, ReadModuleCode((object)codeModule!)));
                    }
                    finally
                    {
                        AccessComAutomation.ReleaseComObject(codeModule);
                        AccessComAutomation.ReleaseComObject(component);
                    }
                }

                modules.Sort((left, right) => string.Compare(left.Name, right.Name, StringComparison.OrdinalIgnoreCase));
                var queries = includeQueries ? ReadQueries((object)app, queryName) : [];
                var resolvedProjectName = Convert.ToString(project.Name) ?? string.Empty;
                return new DbSnapshotSet(resolvedProjectName, modules, queries);
            }
            finally
            {
                AccessComAutomation.ReleaseComObject(components);
                AccessComAutomation.ReleaseComObject(project);
                AccessComAutomation.ReleaseComObject(vbe);
            }
        }, cancellationToken);
    }

    /// <summary>
    /// Reads every saved query's SQL through DAO. Includes Access's own ~ temporary queries and MSys*
    /// system queries, matching what CurrentDb().QueryDefs actually contains.
    /// </summary>
    private static List<QuerySnapshot> ReadQueries(object app, string? queryName)
    {
        object? db = null;
        object? queryDefs = null;
        try
        {
            db = AccessComAutomation.InvokeMethod(app, "CurrentDb");
            queryDefs = db is null ? null : AccessComAutomation.GetProperty(db, "QueryDefs");

            var queries = new List<QuerySnapshot>();
            if (queryDefs is null)
            {
                return queries;
            }

            var count = AccessComAutomation.GetCount(queryDefs);
            // DAO collections are 0-based (unlike VBIDE's 1-based collections).
            for (var i = 0; i < count; i++)
            {
                object? queryDef = null;
                try
                {
                    queryDef = AccessComAutomation.GetItem(queryDefs, i);
                    var name = Convert.ToString(AccessComAutomation.GetProperty(queryDef!, "Name")) ?? string.Empty;
                    if (string.IsNullOrWhiteSpace(name))
                    {
                        continue;
                    }

                    if (!string.IsNullOrWhiteSpace(queryName)
                        && !name.Equals(queryName, StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    var type = 0;
                    var typeValue = AccessComAutomation.TryGetProperty(queryDef!, "Type");
                    if (typeValue is not null)
                    {
                        try
                        {
                            type = Convert.ToInt32(typeValue, CultureInfo.InvariantCulture);
                        }
                        catch
                        {
                        }
                    }

                    var sql = Convert.ToString(AccessComAutomation.GetProperty(queryDef!, "SQL")) ?? string.Empty;
                    queries.Add(new QuerySnapshot(name, type, sql));
                }
                finally
                {
                    AccessComAutomation.ReleaseComObject(queryDef);
                }
            }

            queries.Sort((left, right) => string.Compare(left.Name, right.Name, StringComparison.OrdinalIgnoreCase));
            return queries;
        }
        finally
        {
            AccessComAutomation.ReleaseComObject(queryDefs);
            AccessComAutomation.ReleaseComObject(db);
        }
    }

    [McpServerTool(Name = "access_list_module_procedures", Title = "Access: List Module Procedures", ReadOnly = true, Idempotent = true)]
    [Description("Lists procedures in a VBA module via COM automation.")]
    public async Task<string> ListModuleProceduresAsync(
        [Description("Module name.")] string moduleName,
        [Description("Path to .mdb/.accdb. If omitted, uses DefaultDatabasePath.")] string? databasePath = null,
        CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(moduleName);

        var dbPath = AccessDbPathPolicy.NormalizeAndValidate(databasePath, _options);
        var com = new AccessComAutomation(_options);

        var procedures = await com.RunInAccessAsync(dbPath, app =>
        {
            var items = new List<object>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            // ProcOfLine takes a ByRef argument, so this one stays on the dynamic binder. That is safe
            // here: the crash-prone type libraries are the per-form/report ones, not Module/CodeModule.
            dynamic? codeModule = null;

            try
            {
                codeModule = ResolveCodeModule(app, moduleName);
                var countOfLines = (int)codeModule.CountOfLines;

                var line = 1;
                while (line <= countOfLines)
                {
                    int kind;
                    string? procName = null;
                    try
                    {
                        procName = Convert.ToString(codeModule.ProcOfLine(line, out kind));
                    }
                    catch
                    {
                        line++;
                        continue;
                    }

                    if (string.IsNullOrWhiteSpace(procName))
                    {
                        line++;
                        continue;
                    }

                    var key = $"{procName}:{kind}";
                    if (seen.Add(key))
                    {
                        var startLine = (int)codeModule.ProcStartLine(procName, kind);
                        var lineCount = (int)codeModule.ProcCountLines(procName, kind);
                        items.Add(new
                        {
                            name = procName,
                            kind,
                            startLine,
                            lineCount
                        });
                        line = startLine + Math.Max(1, lineCount);
                    }
                    else
                    {
                        line++;
                    }
                }
            }
            finally
            {
                AccessComAutomation.ReleaseComObject(codeModule);
            }

            return (IReadOnlyList<object>)items;
        }, cancellationToken);

        return JsonSerializer.Serialize(new { databasePath = dbPath, moduleName, procedures }, JsonOptions);
    }

    [McpServerTool(Name = "access_get_module_code", Title = "Access: Get Module Code", ReadOnly = true, Idempotent = true)]
    [Description("Gets the full code for a VBA module via COM automation.")]
    public async Task<string> GetModuleCodeAsync(
        [Description("Module name.")] string moduleName,
        [Description("Path to .mdb/.accdb. If omitted, uses DefaultDatabasePath.")] string? databasePath = null,
        CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(moduleName);

        var dbPath = AccessDbPathPolicy.NormalizeAndValidate(databasePath, _options);
        var com = new AccessComAutomation(_options);

        var code = await com.RunInAccessAsync(dbPath, app =>
        {
            object? codeModule = null;
            try
            {
                codeModule = ResolveCodeModule(app, moduleName);
                var countOfLines = Convert.ToInt32(AccessComAutomation.GetProperty(codeModule, "CountOfLines"), CultureInfo.InvariantCulture);
                return countOfLines > 0
                    ? Convert.ToString(AccessComAutomation.GetProperty(codeModule, "Lines", 1, countOfLines)) ?? string.Empty
                    : string.Empty;
            }
            finally
            {
                AccessComAutomation.ReleaseComObject(codeModule);
            }
        }, cancellationToken);

        return JsonSerializer.Serialize(new { databasePath = dbPath, moduleName, code }, JsonOptions);
    }

    [McpServerTool(Name = "access_export_module_code", Title = "Access: Export Module Code", ReadOnly = true, Idempotent = true)]
    [Description("Exports a VBA module's code as text via COM automation.")]
    public async Task<string> ExportModuleCodeAsync(
        [Description("Module name.")] string moduleName,
        [Description("Path to .mdb/.accdb. If omitted, uses DefaultDatabasePath.")] string? databasePath = null,
        CancellationToken cancellationToken = default)
    {
        return await GetModuleCodeAsync(moduleName, databasePath, cancellationToken);
    }

    [McpServerTool(Name = "access_export_all_modules_to_folder", Title = "Access: Export All Modules To Folder", ReadOnly = true, Idempotent = false)]
    [Description("Exports every VBA module (including code-behind modules for forms and reports) to text files in a folder via VBComponent.Export - so the hidden Attribute VB_Name= and other header attributes are preserved (.bas for standard modules, .frm for form modules, .cls for everything else) - plus a manifest, for git-friendly version control. Files edited locally since the last export are left alone and reported unless overwriteLocalChanges is true; push them back first with access_import_changed_modules_from_folder. With includeQueries, every saved query is also written as Queries/<name>.sql; those files are snapshots only - the drift and import tools cover VBA modules and ignore them.")]
    public async Task<string> ExportAllModulesToFolderAsync(
        [Description("Folder to write exported module files into. Created if it doesn't exist.")] string folderPath,
        [Description("If true, files with local edits since the last export are overwritten. Default false: they're preserved and reported.")] bool overwriteLocalChanges = false,
        [Description("Optional VBA project name. If omitted, uses the first project.")] string? projectName = null,
        [Description("Path to .mdb/.accdb. If omitted, uses DefaultDatabasePath.")] string? databasePath = null,
        [Description("If true, also exports every saved query (CurrentDb().QueryDefs, including Access's ~ temporary queries and MSys* system queries) as Queries/<name>.sql. Default false: existing query files and manifest entries are left exactly as they were.")] bool includeQueries = false,
        CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(folderPath);

        var dbPath = AccessDbPathPolicy.NormalizeAndValidate(databasePath, _options);
        var com = new AccessComAutomation(_options);

        var fullFolder = Path.GetFullPath(folderPath);
        Directory.CreateDirectory(fullFolder);

        // Anything the user edited on disk since the last export is a local change that hasn't been
        // pushed back yet - overwriting it here would silently throw that work away.
        var previousManifest = ReadManifest(fullFolder);
        var previousEntries = (previousManifest?.Modules ?? []).ToDictionary(e => e.Name, StringComparer.OrdinalIgnoreCase);
        var previousQueries = previousManifest?.Queries;
        var previousQueriesByName = (previousQueries ?? []).ToDictionary(e => e.Name, StringComparer.OrdinalIgnoreCase);
        var locallyModified = FindLocallyModified(fullFolder, previousEntries.Values);
        var locallyModifiedQueries = includeQueries && previousQueries is not null
            ? FindLocallyModified(fullFolder, previousQueries)
            : [];

        var result = await com.RunInAccessAsync(dbPath, app =>
        {
            dynamic? vbe = null;
            dynamic? project = null;
            dynamic? components = null;

            try
            {
                vbe = app.VBE;
                project = GetVbeProject(vbe, projectName);
                string resolvedProjectName = Convert.ToString(project.Name) ?? string.Empty;
                components = project.VBComponents;

                var count = (int)components.Count;
                var entries = new List<ExportManifestEntry>();
                var preserved = new List<string>();
                var linesOnly = new List<string>();
                var notRoundTripSafe = new List<string>();

                for (var i = 1; i <= count; i++)
                {
                    var component = components.Item(i);
                    dynamic? codeModule = null;
                    try
                    {
                        string name = Convert.ToString(component.Name) ?? $"Module{i}";
                        int type = Convert.ToInt32(component.Type, CultureInfo.InvariantCulture);
                        codeModule = component.CodeModule;
                        var body = ReadModuleCode((object)codeModule!);
                        var fileName = GetFileName(name, type);
                        var filePath = Path.Combine(fullFolder, fileName);

                        if (!overwriteLocalChanges && locallyModified.Contains(name))
                        {
                            // Keep the old entry too: its hash is the base the user's edits were made
                            // against, which is what the import's staleness check needs to compare to.
                            preserved.Add(name);
                            entries.Add(previousEntries[name]);
                            continue;
                        }

                        var (exportMethod, roundTripSafe) = WriteExportedModuleFile((object)component, body, filePath);
                        if (exportMethod == "lines")
                        {
                            linesOnly.Add(name);
                        }

                        if (!roundTripSafe)
                        {
                            notRoundTripSafe.Add(name);
                        }

                        entries.Add(new ExportManifestEntry(name, type, fileName, exportMethod, roundTripSafe, Sha256Hex(body)));
                    }
                    finally
                    {
                        AccessComAutomation.ReleaseComObject(codeModule);
                        AccessComAutomation.ReleaseComObject(component);
                    }
                }

                var queryEntries = new List<ExportManifestEntry>();
                var preservedQueries = new List<string>();
                if (includeQueries)
                {
                    var usedFileNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    foreach (var query in ReadQueries((object)app, null))
                    {
                        string fileName;
                        if (previousQueriesByName.TryGetValue(query.Name, out var previousQuery)
                            && usedFileNames.Add(previousQuery.FileName))
                        {
                            fileName = previousQuery.FileName;
                        }
                        else
                        {
                            fileName = GetUniqueQueryFileName(query.Name, usedFileNames);
                        }

                        if (!overwriteLocalChanges && locallyModifiedQueries.Contains(query.Name))
                        {
                            // Same base-preservation rule as modules: the old entry is what a later
                            // comparison against the file must be measured from.
                            preservedQueries.Add(query.Name);
                            queryEntries.Add(previousQueriesByName[query.Name]);
                            continue;
                        }

                        WriteQueryFile(Path.Combine(fullFolder, fileName), query.Sql);
                        queryEntries.Add(new ExportManifestEntry(
                            query.Name,
                            query.QueryType,
                            fileName,
                            "sql",
                            RoundTripSafe: true,
                            Sha256Hex(NormalizeQuerySql(query.Sql))));
                    }
                }

                return (projectName: resolvedProjectName, entries, preserved, linesOnly, notRoundTripSafe, queryEntries, preservedQueries);
            }
            finally
            {
                AccessComAutomation.ReleaseComObject(components);
                AccessComAutomation.ReleaseComObject(project);
                AccessComAutomation.ReleaseComObject(vbe);
            }
        }, cancellationToken);

        // Turning includeQueries off must not silently untrack queries from an earlier export.
        var queryEntries = includeQueries ? result.queryEntries : previousQueries;
        var manifestPath = WriteManifest(fullFolder, new ExportManifest(
            dbPath,
            result.projectName,
            DateTimeOffset.UtcNow,
            result.entries,
            queryEntries));

        object? queries = includeQueries
            ? new
            {
                exportedCount = result.queryEntries.Count - result.preservedQueries.Count,
                preservedLocalChanges = result.preservedQueries
            }
            : null;

        return JsonSerializer.Serialize(new
        {
            databasePath = dbPath,
            folderPath = fullFolder,
            projectName = result.projectName,
            exportedCount = result.entries.Count - result.preserved.Count,
            preservedLocalChanges = result.preserved,
            exportedAsLinesOnly = result.linesOnly,
            notRoundTripSafe = result.notRoundTripSafe,
            manifestPath,
            includeQueries,
            queries
        }, OmitNullsJsonOptions);
    }

    [McpServerTool(Name = "access_import_changed_modules_from_folder", Title = "Access: Import Changed Modules From Folder", ReadOnly = false, Idempotent = false)]
    [Description("Pushes back every module whose exported file was edited since access_export_all_modules_to_folder wrote it, per that export's manifest. A module that changed inside Access since the export is refused as stale rather than overwritten; every write is backed up, verified, and rolled back on failure exactly like access_import_module_code. Unchanged files are skipped without touching Access at all.")]
    public async Task<string> ImportChangedModulesFromFolderAsync(
        [Description("Folder previously written by access_export_all_modules_to_folder.")] string folderPath,
        [Description("If true, compiles and saves once after all imports.")] bool compile = true,
        [Description("Path to .mdb/.accdb. If omitted, uses DefaultDatabasePath.")] string? databasePath = null,
        CancellationToken cancellationToken = default)
    {
        if (_options.ReadOnly)
        {
            throw new InvalidOperationException("Server is configured as read-only.");
        }

        ArgumentException.ThrowIfNullOrWhiteSpace(folderPath);

        var dbPath = AccessDbPathPolicy.NormalizeAndValidate(databasePath, _options);
        var fullFolder = Path.GetFullPath(folderPath);
        var manifest = ReadManifest(fullFolder)
            ?? throw new InvalidOperationException($"No {ExportManifestFileName} in '{fullFolder}' - run access_export_all_modules_to_folder first.");

        // Everything about what to push is decided from disk, before Access is involved at all.
        var changed = new List<(ExportManifestEntry Entry, string NewBody)>();
        var skippedNotRoundTripSafe = new List<string>();
        var missingFiles = new List<string>();
        var unchangedCount = 0;

        foreach (var entry in manifest.Modules)
        {
            var filePath = Path.Combine(fullFolder, entry.FileName);
            if (!File.Exists(filePath))
            {
                missingFiles.Add(entry.Name);
                continue;
            }

            var body = StripExportHeaders(File.ReadAllText(filePath));
            if (HashesEqual(Sha256Hex(body), entry.BodySha256))
            {
                unchangedCount++;
                continue;
            }

            if (!entry.RoundTripSafe)
            {
                skippedNotRoundTripSafe.Add(entry.Name);
                continue;
            }

            changed.Add((entry, body));
        }

        var updated = new List<object>();
        var stale = new List<string>();
        var failed = new List<object>();
        var compiled = false;

        if (changed.Count > 0)
        {
            var com = new AccessComAutomation(_options);
            var newHashes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            compiled = await com.RunInAccessAsync(dbPath, app =>
            {
                foreach (var (entry, newBody) in changed)
                {
                    try
                    {
                        // (object) cast: with a dynamic argument the call itself binds dynamically and
                        // its tuple return type is lost, so this would fail to deconstruct.
                        var (lines, backupPath, writtenCode) = ImportModuleCore((object)app, dbPath, entry.Name, newBody, expectedPreviousCode: null, expectedPreviousCodeSha256: entry.BodySha256);
                        updated.Add(new { module = entry.Name, lines, backupPath });
                        newHashes[entry.Name] = Sha256Hex(writtenCode);
                    }
                    catch (StaleModuleException)
                    {
                        stale.Add(entry.Name);
                    }
                    catch (Exception ex)
                    {
                        failed.Add(new { module = entry.Name, error = ex.Message });
                    }
                }

                if (!compile || updated.Count == 0)
                {
                    return false;
                }

                dynamic? doCmd = null;
                try
                {
                    doCmd = app.DoCmd;
                    doCmd.RunCommand(AccessComAutomation.CompileAndSaveAllModulesCommand);
                    return true;
                }
                finally
                {
                    AccessComAutomation.ReleaseComObject(doCmd);
                }
            }, cancellationToken);

            if (newHashes.Count > 0)
            {
                var refreshed = manifest.Modules
                    .Select(e => newHashes.TryGetValue(e.Name, out var hash) ? e with { BodySha256 = hash } : e)
                    .ToList();
                WriteManifest(fullFolder, manifest with { ExportedAtUtc = DateTimeOffset.UtcNow, Modules = refreshed });
            }
        }

        return JsonSerializer.Serialize(new
        {
            databasePath = dbPath,
            folderPath = fullFolder,
            unchangedCount,
            updated,
            stale,
            failed,
            skippedNotRoundTripSafe,
            missingFiles,
            compiled
        }, JsonOptions);
    }

    [McpServerTool(Name = "access_import_module_code", Title = "Access: Import Module Code", ReadOnly = false, Idempotent = false)]
    [Description("Replaces a VBA module's code from text via COM automation. Backs up the module's previous code before writing, and - if expectedPreviousCode is given - refuses to write when the module has changed since it was last read.")]
    public async Task<string> ImportModuleCodeAsync(
        [Description("Module name.")] string moduleName,
        [Description("Module code text to set.")] string code,
        [Description("If true, compiles and saves after import.")] bool compile = true,
        [Description("If provided, the write is refused unless the module's current live code matches this exactly - a guard against clobbering a change someone else made since this was last read.")] string? expectedPreviousCode = null,
        [Description("Like expectedPreviousCode, but as a SHA-256 hex digest of the expected code (UTF-8) - for callers that recorded a hash rather than the full text.")] string? expectedPreviousCodeSha256 = null,
        [Description("Path to .mdb/.accdb. If omitted, uses DefaultDatabasePath.")] string? databasePath = null,
        CancellationToken cancellationToken = default)
    {
        if (_options.ReadOnly)
        {
            throw new InvalidOperationException("Server is configured as read-only.");
        }

        ArgumentException.ThrowIfNullOrWhiteSpace(moduleName);

        var dbPath = AccessDbPathPolicy.NormalizeAndValidate(databasePath, _options);
        var com = new AccessComAutomation(_options);

        var result = await com.RunInAccessAsync(dbPath, app =>
        {
            var imported = ImportModuleCore((object)app, dbPath, moduleName, code, expectedPreviousCode, expectedPreviousCodeSha256);

            if (compile)
            {
                dynamic? doCmd = null;
                try
                {
                    doCmd = app.DoCmd;
                    doCmd.RunCommand(AccessComAutomation.CompileAndSaveAllModulesCommand);
                }
                finally
                {
                    AccessComAutomation.ReleaseComObject(doCmd);
                }
            }

            return imported;
        }, cancellationToken);

        return JsonSerializer.Serialize(new { databasePath = dbPath, moduleName, lines = result.lines, backupPath = result.backupPath }, JsonOptions);
    }

    [McpServerTool(Name = "access_check_export_drift", Title = "Access: Check Export Drift", ReadOnly = true, Idempotent = true)]
    [Description("Compares every VBA module in Access against an export folder's manifest and files without changing anything: which modules changed inside Access since the export (drifted), which were edited locally, which changed in both places (conflicts), plus modules new in or removed from Access. Run it before editing exported files and before importing, to see what moved underneath you.")]
    public async Task<string> CheckExportDriftAsync(
        [Description("Folder previously written by access_export_all_modules_to_folder.")] string folderPath,
        [Description("Optional VBA project name. If omitted, uses the first project.")] string? projectName = null,
        [Description("Path to .mdb/.accdb. If omitted, uses DefaultDatabasePath.")] string? databasePath = null,
        CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(folderPath);

        var dbPath = AccessDbPathPolicy.NormalizeAndValidate(databasePath, _options);
        var fullFolder = Path.GetFullPath(folderPath);
        var manifest = ReadManifest(fullFolder)
            ?? throw new InvalidOperationException($"No {ExportManifestFileName} in '{fullFolder}' - run access_export_all_modules_to_folder first.");

        var com = new AccessComAutomation(_options);
        var live = await com.RunInAccessAsync(dbPath, app => ReadLiveModules((object)app, projectName), cancellationToken);

        var report = ClassifyDrift(fullFolder, manifest, live);
        return JsonSerializer.Serialize(new
        {
            databasePath = dbPath,
            folderPath = fullFolder,
            report.UnchangedCount,
            report.DriftedInAccess,
            report.EditedLocally,
            report.LocalMatchesLive,
            report.Conflicts,
            report.NewInAccess,
            report.RemovedFromAccess,
            report.MissingFiles
        }, JsonOptions);
    }

    [McpServerTool(Name = "access_pull_drifted_modules", Title = "Access: Pull Drifted Modules", ReadOnly = true, Idempotent = false)]
    [Description("Brings an export folder up to date with what changed inside Access since the export, without modifying Access: re-exports modules that drifted in Access but carry no local edits (and modules new in Access), updating their manifest base; for modules changed in BOTH places it writes the live version alongside the local file as <file>.live for merging by hand, leaving the local file and its manifest base untouched so a later import still refuses to overwrite. After merging, pass the module in markResolved to adopt the live version as its base without touching the file, so import accepts the merged result.")]
    public async Task<string> PullDriftedModulesAsync(
        [Description("Folder previously written by access_export_all_modules_to_folder.")] string folderPath,
        [Description("Modules whose local file has already been merged by hand: their manifest base becomes the current live version, the file is left alone, and any <file>.live copy is removed.")] string[]? markResolved = null,
        [Description("Optional VBA project name. If omitted, uses the first project.")] string? projectName = null,
        [Description("Path to .mdb/.accdb. If omitted, uses DefaultDatabasePath.")] string? databasePath = null,
        CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(folderPath);

        var dbPath = AccessDbPathPolicy.NormalizeAndValidate(databasePath, _options);
        var fullFolder = Path.GetFullPath(folderPath);
        var manifest = ReadManifest(fullFolder)
            ?? throw new InvalidOperationException($"No {ExportManifestFileName} in '{fullFolder}' - run access_export_all_modules_to_folder first.");

        var entriesByName = manifest.Modules.ToDictionary(e => e.Name, StringComparer.OrdinalIgnoreCase);
        var localHashes = manifest.Modules.ToDictionary(e => e.Name, e => LocalBodyHash(fullFolder, e), StringComparer.OrdinalIgnoreCase);
        var toResolve = new HashSet<string>(markResolved ?? [], StringComparer.OrdinalIgnoreCase);

        var com = new AccessComAutomation(_options);
        var result = await com.RunInAccessAsync(dbPath, app =>
        {
            dynamic? vbe = null;
            dynamic? project = null;
            dynamic? components = null;

            try
            {
                vbe = app.VBE;
                project = GetVbeProject(vbe, projectName);
                components = project.VBComponents;

                var refreshed = new Dictionary<string, ExportManifestEntry>(StringComparer.OrdinalIgnoreCase);
                var present = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                var newEntries = new List<ExportManifestEntry>();
                var updated = new List<string>();
                var newExported = new List<string>();
                var restoredMissing = new List<string>();
                var resolvedAutomatically = new List<string>();
                var markedResolved = new List<string>();
                var conflicts = new List<object>();
                var editedLocally = new List<string>();
                var unchangedCount = 0;

                var count = (int)components.Count;
                for (var i = 1; i <= count; i++)
                {
                    var component = components.Item(i);
                    dynamic? codeModule = null;
                    try
                    {
                        string name = Convert.ToString(component.Name) ?? $"Module{i}";
                        int type = Convert.ToInt32(component.Type, CultureInfo.InvariantCulture);
                        codeModule = component.CodeModule;
                        var body = ReadModuleCode((object)codeModule!);
                        var liveHash = Sha256Hex(body);

                        if (!entriesByName.TryGetValue(name, out var entry))
                        {
                            var newFileName = GetFileName(name, type);
                            var (method, safe) = WriteExportedModuleFile((object)component, body, Path.Combine(fullFolder, newFileName));
                            newEntries.Add(new ExportManifestEntry(name, type, newFileName, method, safe, liveHash));
                            newExported.Add(name);
                            continue;
                        }

                        present.Add(name);
                        var filePath = Path.Combine(fullFolder, entry.FileName);
                        var sidecarPath = filePath + ".live";
                        var localHash = localHashes[name];
                        var liveChanged = !HashesEqual(liveHash, entry.BodySha256);
                        var localChanged = localHash is not null && !HashesEqual(localHash, entry.BodySha256);

                        if (toResolve.Contains(name))
                        {
                            refreshed[name] = entry with { BodySha256 = liveHash };
                            File.Delete(sidecarPath);
                            markedResolved.Add(name);
                        }
                        else if (!liveChanged)
                        {
                            if (localHash is null)
                            {
                                var (method, safe) = WriteExportedModuleFile((object)component, body, filePath);
                                refreshed[name] = entry with { ExportMethod = method, RoundTripSafe = safe };
                                restoredMissing.Add(name);
                            }
                            else if (localChanged)
                            {
                                editedLocally.Add(name);
                            }
                            else
                            {
                                unchangedCount++;
                            }
                        }
                        else if (localHash is null || !localChanged)
                        {
                            var (method, safe) = WriteExportedModuleFile((object)component, body, filePath);
                            refreshed[name] = entry with { ExportMethod = method, RoundTripSafe = safe, BodySha256 = liveHash };
                            File.Delete(sidecarPath);
                            updated.Add(name);
                        }
                        else if (HashesEqual(localHash, liveHash))
                        {
                            // The local file already matches what's in Access (merged by hand, or the
                            // same change made in both places) - just adopt it as the base.
                            refreshed[name] = entry with { BodySha256 = liveHash };
                            File.Delete(sidecarPath);
                            resolvedAutomatically.Add(name);
                        }
                        else
                        {
                            WriteExportedModuleFile((object)component, body, sidecarPath);
                            conflicts.Add(new { module = name, liveCopy = sidecarPath, localFile = filePath });
                        }
                    }
                    finally
                    {
                        AccessComAutomation.ReleaseComObject(codeModule);
                        AccessComAutomation.ReleaseComObject(component);
                    }
                }

                return (refreshed, present, newEntries, updated, newExported, restoredMissing, resolvedAutomatically, markedResolved, conflicts, editedLocally, unchangedCount);
            }
            finally
            {
                AccessComAutomation.ReleaseComObject(components);
                AccessComAutomation.ReleaseComObject(project);
                AccessComAutomation.ReleaseComObject(vbe);
            }
        }, cancellationToken);

        // Modules the manifest knows but Access no longer has drop out of the manifest; their files are
        // left on disk for the user to deal with rather than deleted.
        var removedFromAccess = new List<string>();
        var kept = new List<ExportManifestEntry>();
        foreach (var entry in manifest.Modules)
        {
            if (result.refreshed.TryGetValue(entry.Name, out var replacement))
            {
                kept.Add(replacement);
            }
            else if (result.present.Contains(entry.Name))
            {
                kept.Add(entry);
            }
            else
            {
                removedFromAccess.Add(entry.Name);
            }
        }

        var finalEntries = kept.Concat(result.newEntries).ToList();
        WriteManifest(fullFolder, manifest with { ExportedAtUtc = DateTimeOffset.UtcNow, Modules = finalEntries });

        return JsonSerializer.Serialize(new
        {
            databasePath = dbPath,
            folderPath = fullFolder,
            result.updated,
            result.newExported,
            result.restoredMissing,
            result.resolvedAutomatically,
            result.markedResolved,
            result.conflicts,
            result.editedLocally,
            removedFromAccess,
            result.unchangedCount
        }, JsonOptions);
    }

    // internal, not private: RunInAccessAsync invokes its callback through dynamic dispatch, and the
    // runtime binder rejects a Func<> whose result type involves a private nested type.
    internal sealed record LiveModule(string Name, int Type, string BodySha256);

    private static List<LiveModule> ReadLiveModules(dynamic app, string? projectName)
    {
        dynamic? vbe = null;
        dynamic? project = null;
        dynamic? components = null;
        try
        {
            vbe = app.VBE;
            project = GetVbeProject(vbe, projectName);
            components = project.VBComponents;

            var modules = new List<LiveModule>();
            var count = (int)components.Count;
            for (var i = 1; i <= count; i++)
            {
                var component = components.Item(i);
                dynamic? codeModule = null;
                try
                {
                    string name = Convert.ToString(component.Name) ?? $"Module{i}";
                    var type = Convert.ToInt32(component.Type, CultureInfo.InvariantCulture);
                    codeModule = component.CodeModule;
                    modules.Add(new LiveModule(name, type, Sha256Hex(ReadModuleCode((object)codeModule!))));
                }
                finally
                {
                    AccessComAutomation.ReleaseComObject(codeModule);
                    AccessComAutomation.ReleaseComObject(component);
                }
            }

            return modules;
        }
        finally
        {
            AccessComAutomation.ReleaseComObject(components);
            AccessComAutomation.ReleaseComObject(project);
            AccessComAutomation.ReleaseComObject(vbe);
        }
    }

    internal sealed class DriftReport
    {
        public int UnchangedCount { get; set; }
        public List<string> DriftedInAccess { get; } = [];
        public List<string> EditedLocally { get; } = [];
        public List<string> LocalMatchesLive { get; } = [];
        public List<string> Conflicts { get; } = [];
        public List<string> NewInAccess { get; } = [];
        public List<string> RemovedFromAccess { get; } = [];
        public List<string> MissingFiles { get; } = [];
    }

    internal static DriftReport ClassifyDrift(string folder, ExportManifest manifest, IReadOnlyList<LiveModule> live)
    {
        var report = new DriftReport();
        var liveByName = live.ToDictionary(m => m.Name, StringComparer.OrdinalIgnoreCase);
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var entry in manifest.Modules)
        {
            seen.Add(entry.Name);
            if (!liveByName.TryGetValue(entry.Name, out var liveModule))
            {
                report.RemovedFromAccess.Add(entry.Name);
                continue;
            }

            var liveChanged = !HashesEqual(liveModule.BodySha256, entry.BodySha256);
            var localHash = LocalBodyHash(folder, entry);
            if (localHash is null)
            {
                report.MissingFiles.Add(entry.Name);
                if (liveChanged)
                {
                    report.DriftedInAccess.Add(entry.Name);
                }

                continue;
            }

            var localChanged = !HashesEqual(localHash, entry.BodySha256);
            switch (liveChanged, localChanged)
            {
                case (false, false):
                    report.UnchangedCount++;
                    break;
                case (true, false):
                    report.DriftedInAccess.Add(entry.Name);
                    break;
                case (false, true):
                    report.EditedLocally.Add(entry.Name);
                    break;
                default:
                    (HashesEqual(localHash, liveModule.BodySha256) ? report.LocalMatchesLive : report.Conflicts).Add(entry.Name);
                    break;
            }
        }

        foreach (var module in live)
        {
            if (!seen.Contains(module.Name))
            {
                report.NewInAccess.Add(module.Name);
            }
        }

        return report;
    }

    private static string? LocalBodyHash(string folder, ExportManifestEntry entry)
    {
        var path = Path.Combine(folder, entry.FileName);
        return File.Exists(path) ? Sha256Hex(ReadExportedBody(path, entry)) : null;
    }

    /// <summary>
    /// Names whose exported file no longer matches the body hash recorded at export time - local edits
    /// that must not be overwritten by the next export.
    /// </summary>
    private static HashSet<string> FindLocallyModified(string folder, IEnumerable<ExportManifestEntry> entries)
    {
        var modified = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var entry in entries)
        {
            var path = Path.Combine(folder, entry.FileName);
            if (File.Exists(path) && !HashesEqual(Sha256Hex(ReadExportedBody(path, entry)), entry.BodySha256))
            {
                modified.Add(entry.Name);
            }
        }

        return modified;
    }

    /// <summary>Reads an exported file back to the exact body its manifest hash was computed over.</summary>
    private static string ReadExportedBody(string filePath, ExportManifestEntry entry)
        => entry.ExportMethod == "sql"
            ? ReadExportedQueryBody(File.ReadAllText(filePath))
            : StripExportHeaders(File.ReadAllText(filePath));

    /// <summary>
    /// Writes one module's export file: VBComponent.Export re-saved as UTF-8 (VBE writes the ANSI code
    /// page, which git and most editors would misread), falling back to the bare code body if Export
    /// throws. Reports whether stripping the headers back off reproduces the code body exactly - the
    /// precondition for ever importing the file back.
    /// </summary>
    private static (string exportMethod, bool roundTripSafe) WriteExportedModuleFile(dynamic component, string body, string filePath)
    {
        string exportMethod;
        try
        {
            if (File.Exists(filePath))
            {
                File.Delete(filePath);
            }
            else
                Directory.CreateDirectory(
                    Path.GetDirectoryName(filePath) ?? throw new InvalidOperationException($"Cannot determine directory for '{filePath}'")
                );

            component.Export(filePath);
            var exportedText = File.ReadAllText(filePath, AnsiEncoding.Value);
            File.WriteAllText(filePath, exportedText, Utf8NoBom);
            exportMethod = "export";
        }
        catch (Exception)
        {
            File.WriteAllText(filePath, body + "\r\n", Utf8NoBom);
            exportMethod = "lines";
        }

        var roundTripSafe = string.Equals(StripExportHeaders(File.ReadAllText(filePath)), body, StringComparison.Ordinal);
        return (exportMethod, roundTripSafe);
    }

    /// <summary>
    /// Thrown when a write is refused because the module changed inside Access since the caller last
    /// read it - distinct from a write that was attempted and failed, so batch callers can report the
    /// two differently.
    /// </summary>
    internal sealed class StaleModuleException(string message) : InvalidOperationException(message);

    /// <summary>
    /// The single-module write, against an already-open Access instance, shared by
    /// <see cref="ImportModuleCodeAsync"/> and <see cref="ImportChangedModulesFromFolderAsync"/>.
    /// Backs up, checks staleness, writes, verifies, and rolls back on any failure. Does not compile -
    /// callers decide whether/when to do that, since a batch wants it once at the end.
    /// </summary>
    private static (int lines, string backupPath, string code) ImportModuleCore(
        dynamic app,
        string dbPath,
        string moduleName,
        string? code,
        string? expectedPreviousCode,
        string? expectedPreviousCodeSha256)
    {
        object? codeModule = null;
        try
        {
            codeModule = ResolveCodeModule(app, moduleName);
            var currentCode = ReadModuleCode(codeModule);

            var stale = (expectedPreviousCode is not null && !string.Equals(currentCode, expectedPreviousCode, StringComparison.Ordinal))
                || (expectedPreviousCodeSha256 is not null && !HashesEqual(Sha256Hex(currentCode), expectedPreviousCodeSha256));
            if (stale)
            {
                throw new StaleModuleException(
                    $"Module '{moduleName}' has changed since it was last read - refusing to overwrite. Pull the latest code and retry.");
            }

            var backupPath = BackupModuleCode(dbPath, moduleName, currentCode);
            var intended = VbaProjectEditor.PrepareForWrite(code ?? string.Empty, currentCode);
            VbaProjectEditor.ValidateLineLengths(intended.Code, moduleName);
            var intendedCode = intended.Code;

            // Access's own VBIDE automation has a real, reproducible failure mode where a rewrite
            // either throws partway through (leaving a partial write in place) or completes without
            // error yet silently corrupts content (confirmed independent of this codebase - plain
            // PowerShell COM automation hit the same corruption). Neither is specific to any one API
            // call. A write is never trusted on faith: attempt it, read back, and on ANY failure -
            // thrown exception or content mismatch - try to restore the pre-write content and fail
            // loudly, rather than leave the module in a state nobody asked for.
            Exception? writeFailure = null;
            try
            {
                WriteModuleCode(codeModule, intendedCode);
                var comparison = VbaProjectEditor.CompareCode(intendedCode, ReadModuleCode(codeModule));
                if (!comparison.Same)
                {
                    writeFailure = new InvalidOperationException(
                        "write completed but content did not verify (Access's automation silently altered it) - "
                        + VbaProjectEditor.DescribeMismatch(comparison));
                }
            }
            catch (Exception ex)
            {
                writeFailure = ex;
            }

            if (writeFailure is not null)
            {
                var rolledBack = false;
                try
                {
                    WriteModuleCode(codeModule, currentCode);
                    rolledBack = VbaProjectEditor.CompareCode(currentCode, ReadModuleCode(codeModule)).Same;
                }
                catch
                {
                    rolledBack = false;
                }

                throw new InvalidOperationException(rolledBack
                    ? $"Write to module '{moduleName}' failed ({writeFailure.Message}) and has been rolled back to its previous content. This is a known reliability issue with Access COM automation, not specific to this tool; retry, or edit this module by hand. Previous content is also saved at {backupPath}."
                    : $"Write to module '{moduleName}' failed ({writeFailure.Message}), AND the automatic rollback also failed to verify. The module's content may now be inconsistent - restore it manually from {backupPath} before touching this module again.",
                    writeFailure);
            }

            var newLines = Convert.ToInt32(AccessComAutomation.GetProperty(codeModule, "CountOfLines"), CultureInfo.InvariantCulture);
            return (newLines, backupPath, intendedCode);
        }
        finally
        {
            AccessComAutomation.ReleaseComObject(codeModule);
        }
    }

    private const string ExportManifestFileName = "_accessmcp_manifest.json";

    internal sealed record ExportManifest(
        string DatabasePath,
        string ProjectName,
        DateTimeOffset ExportedAtUtc,
        List<ExportManifestEntry> Modules,
        List<ExportManifestEntry>? Queries = null);

    /// <param name="ExportMethod">"export" (VBComponent.Export, headers preserved) or "lines" (code body only, the fallback when Export throws).</param>
    /// <param name="RoundTripSafe">Whether stripping this file's export headers reproduces the code body exactly - the precondition for importing it back.</param>
    /// <param name="BodySha256">SHA-256 of the code body (UTF-8) as it was in Access at export time - the base for both local-change and staleness detection.</param>
    internal sealed record ExportManifestEntry(
        string Name,
        int Type,
        string FileName,
        string ExportMethod,
        bool RoundTripSafe,
        string BodySha256);

    private static readonly UTF8Encoding Utf8NoBom = new(encoderShouldEmitUTF8Identifier: false);

    private static readonly Lazy<Encoding> AnsiEncoding = new(() =>
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        return Encoding.GetEncoding(CultureInfo.CurrentCulture.TextInfo.ANSICodePage);
    });

    private static ExportManifest? ReadManifest(string folder)
    {
        var path = Path.Combine(folder, ExportManifestFileName);
        if (!File.Exists(path))
        {
            return null;
        }

        try
        {
            return JsonSerializer.Deserialize<ExportManifest>(File.ReadAllText(path), JsonOptions);
        }
        catch (JsonException)
        {
            return null;
        }
    }

    private static string WriteManifest(string folder, ExportManifest manifest)
    {
        var path = Path.Combine(folder, ExportManifestFileName);
        // Omit nulls so manifests written without queries stay identical to the pre-query format.
        File.WriteAllText(path, JsonSerializer.Serialize(manifest, OmitNullsJsonOptions), Utf8NoBom);
        return path;
    }

    /// <summary>
    /// Reduces a VBComponent.Export file to the code body CodeModule.Lines would return: drops the
    /// class-module VERSION/BEGIN..END header block and every "Attribute ..." line (a reserved word
    /// in VBA, so never the start of real code). The file's terminating newline is removed exactly
    /// once so a module whose real last line is blank still round-trips.
    /// </summary>
    internal static string StripExportHeaders(string fileText)
    {
        var lines = fileText.Replace("\r\n", "\n").Split('\n').ToList();
        if (lines.Count > 0 && lines[^1].Length == 0)
        {
            lines.RemoveAt(lines.Count - 1);
        }

        var start = 0;
        if (lines.Count > 0 && lines[0].StartsWith("VERSION ", StringComparison.OrdinalIgnoreCase))
        {
            var depth = 0;
            var i = 1;
            for (; i < lines.Count; i++)
            {
                var trimmed = lines[i].Trim();
                if (trimmed.StartsWith("BEGIN", StringComparison.OrdinalIgnoreCase))
                {
                    depth++;
                }
                else if (trimmed.Equals("END", StringComparison.OrdinalIgnoreCase))
                {
                    depth--;
                    if (depth <= 0)
                    {
                        i++;
                        break;
                    }
                }
            }

            start = i;
        }

        var body = new List<string>();
        for (var i = start; i < lines.Count; i++)
        {
            if (lines[i].TrimStart().StartsWith("Attribute ", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            body.Add(lines[i]);
        }

        return string.Join("\r\n", body);
    }

    private static string Sha256Hex(string text)
        => Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(text)));

    private static bool HashesEqual(string a, string b)
        => string.Equals(a, b, StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// Writes one query's SQL as UTF-8 (no BOM) with CRLF line endings and a single trailing newline, so
    /// the file is git-friendly and <see cref="ReadExportedQueryBody"/> can recover the exact SQL.
    /// </summary>
    private static void WriteQueryFile(string filePath, string sql)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(filePath)
            ?? throw new InvalidOperationException($"Cannot determine directory for '{filePath}'"));
        File.WriteAllText(filePath, NormalizeQuerySql(sql) + "\r\n", Utf8NoBom);
    }

    /// <summary>
    /// Canonical line endings for query SQL, used for both the file body and its manifest hash, so an
    /// editor that saves LF does not read as a local edit.
    /// </summary>
    internal static string NormalizeQuerySql(string sql)
        => sql.Replace("\r\n", "\n").Replace('\r', '\n').Replace("\n", "\r\n");

    /// <summary>Drops the single trailing newline <see cref="WriteQueryFile"/> adds.</summary>
    internal static string ReadExportedQueryBody(string fileText)
    {
        var normalized = NormalizeQuerySql(fileText);
        return normalized.EndsWith("\r\n", StringComparison.Ordinal) ? normalized[..^2] : normalized;
    }

    private static string GetQueryFileName(string name)
        => Path.Combine("Queries", SafeFileName(name) + ".sql");

    /// <summary>
    /// Reserves a unique path for a query file. Two query names can sanitize to the same file name (or a
    /// file name can collide with an earlier one), so a short hash of the exact name is appended when
    /// needed - deterministic across exports, unlike an index.
    /// </summary>
    internal static string GetUniqueQueryFileName(string name, HashSet<string> usedFileNames)
    {
        var candidate = GetQueryFileName(name);
        if (usedFileNames.Add(candidate))
        {
            return candidate;
        }

        var hash = Sha256Hex(name)[..8];
        candidate = Path.Combine("Queries", $"{SafeFileName(name)}~{hash}.sql");
        var suffix = 1;
        while (!usedFileNames.Add(candidate))
        {
            candidate = Path.Combine("Queries", $"{SafeFileName(name)}~{hash}-{suffix++}.sql");
        }

        return candidate;
    }

    private static string GetFileName(string name, int componentType)
        => AddFolder(SafeFileName(name) + ExtensionForComponentType(componentType));

    private static string SafeFileName(string name)
        => new(name.Select(c => Path.GetInvalidFileNameChars().Contains(c) ? '_' : c).ToArray());

    private static string AddFolder(string filename)
    {
        if (filename.StartsWith("Form_", StringComparison.OrdinalIgnoreCase))
        {
            return Path.Combine("Forms", filename);
        }
        if (filename.StartsWith("Report_", StringComparison.OrdinalIgnoreCase))
        {
            return Path.Combine("Reports", filename);
        }
        if (filename.EndsWith(".cls", StringComparison.OrdinalIgnoreCase))
        {
            return Path.Combine("Classes", filename);
        }
        return Path.Combine("Modules", filename);
    }

    private static string ExtensionForComponentType(int componentType) => componentType switch
    {
        1 => ".bas",
        3 => ".frm",
        _ => ".cls",
    };

    private static string ReadModuleCode(object codeModule)
    {
        var countOfLines = Convert.ToInt32(AccessComAutomation.GetProperty(codeModule, "CountOfLines"), CultureInfo.InvariantCulture);
        return countOfLines > 0
            ? Convert.ToString(AccessComAutomation.GetProperty(codeModule, "Lines", 1, countOfLines)) ?? string.Empty
            : string.Empty;
    }

    /// <summary>
    /// Replaces a module's full text as a minimal line-level edit through the shared writer in
    /// <see cref="VbaProjectEditor"/> (ReplaceLine where a line differs, InsertLines only for lines past
    /// the old end, DeleteLines only for surplus old lines; never AddFromString). Measured against a
    /// real database: a single AddFromString call silently corrupts content;
    /// delete-everything-then-InsertLines leaves an extra empty trailing line (69 inserts into an
    /// emptied module read back as 70), which the writer's fixup removes; ReplaceLine, including with
    /// an empty string, is exact.
    /// </summary>
    private static void WriteModuleCode(object codeModule, string code)
        => VbaProjectEditor.WriteCode(codeModule, code);

    /// <summary>
    /// Saves a module's pre-overwrite code to a timestamped file next to the database, so a bad
    /// import can always be undone by hand even without expectedPreviousCode drift protection.
    /// </summary>
    private static string BackupModuleCode(string databasePath, string moduleName, string code)
    {
        var dir = Path.Combine(Path.GetDirectoryName(databasePath) ?? ".", "_accessmcp_backups");
        Directory.CreateDirectory(dir);

        var safeModuleName = new string(moduleName.Select(c => Path.GetInvalidFileNameChars().Contains(c) ? '_' : c).ToArray());
        var timestamp = DateTime.UtcNow.ToString("yyyyMMdd_HHmmssfff", CultureInfo.InvariantCulture);
        var path = Path.Combine(dir, $"{safeModuleName}_{timestamp}.bas");

        File.WriteAllText(path, code);
        return path;
    }

    [McpServerTool(Name = "access_compile_vba", Title = "Access: Compile VBA", ReadOnly = false, Idempotent = false)]
    [Description("Compiles and saves all VBA modules and reports the real result via Application.IsCompiled, with broken-reference preflight and best-effort diagnostics. Always returns the same fields: check 'compiled'; on failure 'module'/'line'/'message' locate a compile error, and 'timedOut'/'crashed'/'cancelled'/'error' describe a session that didn't finish.")]
    public async Task<string> CompileVbaAsync(
        [Description("Path to .mdb/.accdb. If omitted, uses DefaultDatabasePath.")] string? databasePath = null,
        CancellationToken cancellationToken = default)
    {
        if (_options.ReadOnly)
        {
            throw new InvalidOperationException("Server is configured as read-only.");
        }

        var dbPath = AccessDbPathPolicy.NormalizeAndValidate(databasePath, _options);
        var com = new AccessComAutomation(_options);

        var outcome = await com.RunInAccessWithWatchdogAsync(
            dbPath,
            app => VbaCompiler.Compile((object)app),
            new AccessRunOptions
            {
                Timeout = TimeSpan.FromSeconds(Math.Max(1, _options.VbaCompileTimeoutSeconds))
            },
            cancellationToken);

        // One shape whether Access compiled, reported an error, hung or died, so a caller can always
        // read `compiled` and the diagnostics without checking which fields are present.
        var result = outcome.Ok ? outcome.Value : null;
        return JsonSerializer.Serialize(new
        {
            databasePath = dbPath,
            compiled = result?.Ok ?? false,
            isCompiled = result?.IsCompiled ?? false,
            module = result?.Module,
            line = result?.Line,
            message = result?.Message,
            errorNumber = result?.ErrorNumber,
            dialogText = outcome.DialogText ?? result?.DialogText,
            brokenReferenceFlag = result?.BrokenReferenceFlag ?? false,
            brokenReferences = result?.BrokenReferences ?? [],
            timedOut = outcome.TimedOut,
            crashed = outcome.Crashed,
            cancelled = outcome.Cancelled,
            exitCode = outcome.ExitCode,
            error = outcome.Exception?.Message
        }, JsonOptions);
    }

    [McpServerTool(Name = "access_list_reports", Title = "Access: List Reports", ReadOnly = true, Idempotent = true)]
    [Description("Lists Access reports via COM automation.")]
    public async Task<string> ListReportsAsync(
        [Description("Path to .mdb/.accdb. If omitted, uses DefaultDatabasePath.")] string? databasePath = null,
        CancellationToken cancellationToken = default)
    {
        var dbPath = AccessDbPathPolicy.NormalizeAndValidate(databasePath, _options);
        var com = new AccessComAutomation(_options);

        var reports = await com.RunInAccessAsync(dbPath, app =>
        {
            var names = new List<string>();
            dynamic? allReports = null;

            try
            {
                allReports = app.CurrentProject.AllReports;
                var count = (int)allReports.Count;
                for (var i = 0; i < count; i++)
                {
                    var reportRef = allReports.Item(i);
                    try
                    {
                        var name = Convert.ToString(reportRef.Name);
                        if (!string.IsNullOrWhiteSpace(name))
                        {
                            names.Add(name);
                        }
                    }
                    finally
                    {
                        AccessComAutomation.ReleaseComObject(reportRef);
                    }
                }
            }
            finally
            {
                AccessComAutomation.ReleaseComObject(allReports);
            }

            names.Sort(StringComparer.OrdinalIgnoreCase);
            return (IReadOnlyList<string>)names;
        }, cancellationToken);

        return JsonSerializer.Serialize(new { databasePath = dbPath, reports }, JsonOptions);
    }

    [McpServerTool(Name = "access_list_macros", Title = "Access: List Macros", ReadOnly = true, Idempotent = true)]
    [Description("Lists Access macros via COM automation.")]
    public async Task<string> ListMacrosAsync(
        [Description("Path to .mdb/.accdb. If omitted, uses DefaultDatabasePath.")] string? databasePath = null,
        CancellationToken cancellationToken = default)
    {
        var dbPath = AccessDbPathPolicy.NormalizeAndValidate(databasePath, _options);
        var com = new AccessComAutomation(_options);

        var macros = await com.RunInAccessAsync(dbPath, app =>
        {
            var names = new List<string>();
            dynamic? allMacros = null;

            try
            {
                allMacros = app.CurrentProject.AllMacros;
                var count = (int)allMacros.Count;
                for (var i = 0; i < count; i++)
                {
                    var macroRef = allMacros.Item(i);
                    try
                    {
                        var name = Convert.ToString(macroRef.Name);
                        if (!string.IsNullOrWhiteSpace(name))
                        {
                            names.Add(name);
                        }
                    }
                    finally
                    {
                        AccessComAutomation.ReleaseComObject(macroRef);
                    }
                }
            }
            finally
            {
                AccessComAutomation.ReleaseComObject(allMacros);
            }

            names.Sort(StringComparer.OrdinalIgnoreCase);
            return (IReadOnlyList<string>)names;
        }, cancellationToken);

        return JsonSerializer.Serialize(new { databasePath = dbPath, macros }, JsonOptions);
    }

    [McpServerTool(Name = "access_get_table_design", Title = "Access: Get Table Design", ReadOnly = true, Idempotent = true)]
    [Description("Lists tables with field and index design metadata via DAO, without launching Access.")]
    public async Task<string> GetTableDesignAsync(
        [Description("Optional table name filter.")] string? tableName = null,
        [Description("Path to .mdb/.accdb. If omitted, uses DefaultDatabasePath.")] string? databasePath = null,
        CancellationToken cancellationToken = default)
    {
        var dbPath = AccessDbPathPolicy.NormalizeAndValidate(databasePath, _options);

        var dao = new AccessDao(_options);
        var tables = await dao.GetTableDesignAsync(dbPath, tableName, cancellationToken);

        return JsonSerializer.Serialize(new { databasePath = dbPath, tables }, JsonOptions);
    }

    [McpServerTool(Name = "access_get_object_properties", Title = "Access: Get Object Properties", ReadOnly = true, Idempotent = true)]
    [Description("Gets properties for a form or report via COM automation.")]
    public async Task<string> GetObjectPropertiesAsync(
        [Description("Object type: form or report.")] string objectType,
        [Description("Form or report name.")] string objectName,
        [Description("Optional list of property names. If omitted, returns all properties.")] string[]? properties = null,
        [Description("Path to .mdb/.accdb. If omitted, uses DefaultDatabasePath.")] string? databasePath = null,
        CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(objectType);
        ArgumentException.ThrowIfNullOrWhiteSpace(objectName);

        var dbPath = AccessDbPathPolicy.NormalizeAndValidate(databasePath, _options);
        var com = new AccessComAutomation(_options);
        var normalizedType = objectType.Trim().ToLowerInvariant();

        var values = await com.RunInAccessAsync(dbPath, app =>
        {
            var doCmd = app.DoCmd;
            // Held as `object`, not `dynamic` — see the note on AccessComAutomation's reflection helpers.
            object? target = null;

            try
            {
                if (normalizedType == "form")
                {
                    doCmd.OpenForm(objectName, AccessComAutomation.ResolveFormView(true));
                    target = app.Forms[objectName];
                }
                else if (normalizedType == "report")
                {
                    doCmd.OpenReport(objectName, AccessComAutomation.ResolveReportView("design"));
                    target = app.Reports[objectName];
                }
                else
                {
                    throw new ArgumentException("Unsupported object type. Expected 'form' or 'report'.", nameof(objectType));
                }

                var result = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
                if (properties is not null && properties.Length > 0)
                {
                    foreach (var propName in properties.Where(p => !string.IsNullOrWhiteSpace(p)))
                    {
                        result[propName] = AccessComAutomation.TryGetProperty(target!, propName);
                    }
                }
                else
                {
                    object? props = null;
                    try
                    {
                        props = AccessComAutomation.GetProperty(target!, "Properties")!;
                        var count = AccessComAutomation.GetCount(props);
                        for (var i = 0; i < count; i++)
                        {
                            var prop = AccessComAutomation.GetItem(props, i)!;
                            try
                            {
                                var name = Convert.ToString(AccessComAutomation.GetProperty(prop, "Name")) ?? string.Empty;
                                if (!string.IsNullOrWhiteSpace(name))
                                {
                                    result[name] = AccessComAutomation.TryGetProperty(prop, "Value");
                                }
                            }
                            finally
                            {
                                AccessComAutomation.ReleaseComObject(prop);
                            }
                        }
                    }
                    finally
                    {
                        AccessComAutomation.ReleaseComObject(props);
                    }
                }

                return (IReadOnlyDictionary<string, object?>)result;
            }
            finally
            {
                try
                {
                    var closeSave = AccessComAutomation.ResolveCloseSave(false);
                    doCmd.Close(normalizedType == "form" ? AccessComAutomation.FormObjectType : AccessComAutomation.ReportObjectType, objectName, closeSave);
                }
                catch
                {
                }

                AccessComAutomation.ReleaseComObject(target);
                AccessComAutomation.ReleaseComObject(doCmd);
            }
        }, cancellationToken);

        return JsonSerializer.Serialize(new { databasePath = dbPath, objectType = normalizedType, objectName, properties = values }, JsonOptions);
    }

    [McpServerTool(Name = "access_set_object_properties", Title = "Access: Set Object Properties", ReadOnly = false, Idempotent = false)]
    [Description("Sets properties for a form or report via COM automation.")]
    public async Task<string> SetObjectPropertiesAsync(
        [Description("Object type: form or report.")] string objectType,
        [Description("Form or report name.")] string objectName,
        [Description("Properties as JSON object.")] JsonElement? properties,
        [Description("Path to .mdb/.accdb. If omitted, uses DefaultDatabasePath.")] string? databasePath = null,
        CancellationToken cancellationToken = default)
    {
        if (_options.ReadOnly)
        {
            throw new InvalidOperationException("Server is configured as read-only.");
        }

        ArgumentException.ThrowIfNullOrWhiteSpace(objectType);
        ArgumentException.ThrowIfNullOrWhiteSpace(objectName);

        var props = ParseJsonObject(properties);
        if (props is null || props.Count == 0)
        {
            throw new ArgumentException("Properties must be a JSON object with at least one entry.", nameof(properties));
        }

        var dbPath = AccessDbPathPolicy.NormalizeAndValidate(databasePath, _options);
        var com = new AccessComAutomation(_options);
        var normalizedType = objectType.Trim().ToLowerInvariant();

        var result = await com.RunInAccessAsync(dbPath, app =>
        {
            var doCmd = app.DoCmd;
            // Held as `object`, not `dynamic` — see the note on AccessComAutomation's reflection helpers.
            object? target = null;
            var set = new List<string>();
            var failed = new List<string>();

            try
            {
                if (normalizedType == "form")
                {
                    doCmd.OpenForm(objectName, AccessComAutomation.ResolveFormView(true));
                    target = app.Forms[objectName];
                }
                else if (normalizedType == "report")
                {
                    doCmd.OpenReport(objectName, AccessComAutomation.ResolveReportView("design"));
                    target = app.Reports[objectName];
                }
                else
                {
                    throw new ArgumentException("Unsupported object type. Expected 'form' or 'report'.", nameof(objectType));
                }

                foreach (var (name, value) in props)
                {
                    if (AccessComAutomation.TrySetProperty(target!, name, value))
                    {
                        set.Add(name);
                    }
                    else
                    {
                        failed.Add(name);
                    }
                }

                return new { set, failed };
            }
            finally
            {
                try
                {
                    var closeSave = AccessComAutomation.ResolveCloseSave(true);
                    doCmd.Close(normalizedType == "form" ? AccessComAutomation.FormObjectType : AccessComAutomation.ReportObjectType, objectName, closeSave);
                }
                catch
                {
                }

                AccessComAutomation.ReleaseComObject(target);
                AccessComAutomation.ReleaseComObject(doCmd);
            }
        }, cancellationToken);

        return JsonSerializer.Serialize(new { databasePath = dbPath, objectType = normalizedType, objectName, set = result.set, failed = result.failed }, JsonOptions);
    }

    [McpServerTool(Name = "access_get_control_properties", Title = "Access: Get Control Properties", ReadOnly = true, Idempotent = true)]
    [Description("Gets properties for a form or report control via COM automation.")]
    public async Task<string> GetControlPropertiesAsync(
        [Description("Object type: form or report.")] string objectType,
        [Description("Form or report name.")] string objectName,
        [Description("Control name.")] string controlName,
        [Description("Optional list of property names. If omitted, returns all properties.")] string[]? properties = null,
        [Description("If true, opens in view mode to read Value when available.")] bool includeValue = false,
        [Description("Path to .mdb/.accdb. If omitted, uses DefaultDatabasePath.")] string? databasePath = null,
        CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(objectType);
        ArgumentException.ThrowIfNullOrWhiteSpace(objectName);
        ArgumentException.ThrowIfNullOrWhiteSpace(controlName);

        var dbPath = AccessDbPathPolicy.NormalizeAndValidate(databasePath, _options);
        var com = new AccessComAutomation(_options);
        var normalizedType = objectType.Trim().ToLowerInvariant();

        var values = await com.RunInAccessAsync(dbPath, app =>
        {
            var doCmd = app.DoCmd;
            // Held as `object`, not `dynamic` — see the note on AccessComAutomation's reflection helpers.
            object? target = null;
            object? control = null;

            try
            {
                if (normalizedType == "form")
                {
                    doCmd.OpenForm(objectName, AccessComAutomation.ResolveFormView(!includeValue));
                    target = app.Forms[objectName];
                }
                else if (normalizedType == "report")
                {
                    var view = includeValue ? "report" : "design";
                    doCmd.OpenReport(objectName, AccessComAutomation.ResolveReportView(view));
                    target = app.Reports[objectName];
                }
                else
                {
                    throw new ArgumentException("Unsupported object type. Expected 'form' or 'report'.", nameof(objectType));
                }

                control = AccessComAutomation.GetItem(AccessComAutomation.GetProperty(target!, "Controls")!, controlName)!;

                var result = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
                if (properties is not null && properties.Length > 0)
                {
                    foreach (var propName in properties.Where(p => !string.IsNullOrWhiteSpace(p)))
                    {
                        result[propName] = AccessComAutomation.TryGetProperty(control, propName);
                    }
                }
                else
                {
                    object? props = null;
                    try
                    {
                        props = AccessComAutomation.GetProperty(control, "Properties")!;
                        var count = AccessComAutomation.GetCount(props);
                        for (var i = 0; i < count; i++)
                        {
                            var prop = AccessComAutomation.GetItem(props, i)!;
                            try
                            {
                                var name = Convert.ToString(AccessComAutomation.GetProperty(prop, "Name")) ?? string.Empty;
                                if (!string.IsNullOrWhiteSpace(name))
                                {
                                    result[name] = AccessComAutomation.TryGetProperty(prop, "Value");
                                }
                            }
                            finally
                            {
                                AccessComAutomation.ReleaseComObject(prop);
                            }
                        }
                    }
                    finally
                    {
                        AccessComAutomation.ReleaseComObject(props);
                    }
                }

                return (IReadOnlyDictionary<string, object?>)result;
            }
            finally
            {
                try
                {
                    var closeSave = AccessComAutomation.ResolveCloseSave(false);
                    doCmd.Close(normalizedType == "form" ? AccessComAutomation.FormObjectType : AccessComAutomation.ReportObjectType, objectName, closeSave);
                }
                catch
                {
                }

                AccessComAutomation.ReleaseComObject(control);
                AccessComAutomation.ReleaseComObject(target);
                AccessComAutomation.ReleaseComObject(doCmd);
            }
        }, cancellationToken);

        return JsonSerializer.Serialize(new { databasePath = dbPath, objectType = normalizedType, objectName, controlName, properties = values }, JsonOptions);
    }

    [McpServerTool(Name = "access_set_control_properties", Title = "Access: Set Control Properties", ReadOnly = false, Idempotent = false)]
    [Description("Sets properties for a form or report control via COM automation.")]
    public async Task<string> SetControlPropertiesAsync(
        [Description("Object type: form or report.")] string objectType,
        [Description("Form or report name.")] string objectName,
        [Description("Control name.")] string controlName,
        [Description("Properties as JSON object.")] JsonElement? properties,
        [Description("Path to .mdb/.accdb. If omitted, uses DefaultDatabasePath.")] string? databasePath = null,
        CancellationToken cancellationToken = default)
    {
        if (_options.ReadOnly)
        {
            throw new InvalidOperationException("Server is configured as read-only.");
        }

        ArgumentException.ThrowIfNullOrWhiteSpace(objectType);
        ArgumentException.ThrowIfNullOrWhiteSpace(objectName);
        ArgumentException.ThrowIfNullOrWhiteSpace(controlName);

        var props = ParseJsonObject(properties);
        if (props is null || props.Count == 0)
        {
            throw new ArgumentException("Properties must be a JSON object with at least one entry.", nameof(properties));
        }

        var dbPath = AccessDbPathPolicy.NormalizeAndValidate(databasePath, _options);
        var com = new AccessComAutomation(_options);
        var normalizedType = objectType.Trim().ToLowerInvariant();

        var result = await com.RunInAccessAsync(dbPath, app =>
        {
            var doCmd = app.DoCmd;
            // Held as `object`, not `dynamic` — see the note on AccessComAutomation's reflection helpers.
            object? target = null;
            object? control = null;
            var set = new List<string>();
            var failed = new List<string>();

            try
            {
                if (normalizedType == "form")
                {
                    doCmd.OpenForm(objectName, AccessComAutomation.ResolveFormView(true));
                    target = app.Forms[objectName];
                }
                else if (normalizedType == "report")
                {
                    doCmd.OpenReport(objectName, AccessComAutomation.ResolveReportView("design"));
                    target = app.Reports[objectName];
                }
                else
                {
                    throw new ArgumentException("Unsupported object type. Expected 'form' or 'report'.", nameof(objectType));
                }

                control = AccessComAutomation.GetItem(AccessComAutomation.GetProperty(target!, "Controls")!, controlName)!;

                foreach (var (name, value) in props)
                {
                    if (AccessComAutomation.TrySetProperty(control, name, value))
                    {
                        set.Add(name);
                    }
                    else
                    {
                        failed.Add(name);
                    }
                }

                return new { set, failed };
            }
            finally
            {
                try
                {
                    var closeSave = AccessComAutomation.ResolveCloseSave(true);
                    doCmd.Close(normalizedType == "form" ? AccessComAutomation.FormObjectType : AccessComAutomation.ReportObjectType, objectName, closeSave);
                }
                catch
                {
                }

                AccessComAutomation.ReleaseComObject(control);
                AccessComAutomation.ReleaseComObject(target);
                AccessComAutomation.ReleaseComObject(doCmd);
            }
        }, cancellationToken);

        return JsonSerializer.Serialize(new { databasePath = dbPath, objectType = normalizedType, objectName, controlName, set = result.set, failed = result.failed }, JsonOptions);
    }

    [McpServerTool(Name = "access_get_report_controls", Title = "Access: Get Report Controls", ReadOnly = true, Idempotent = true)]
    [Description("Gets controls for an Access report via COM automation.")]
    public async Task<string> GetReportControlsAsync(
        [Description("Report name.")] string reportName,
        [Description("If true, includes current control values (report view).")]
        bool includeValues = false,
        [Description("Path to .mdb/.accdb. If omitted, uses DefaultDatabasePath.")] string? databasePath = null,
        CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(reportName);

        var dbPath = AccessDbPathPolicy.NormalizeAndValidate(databasePath, _options);
        var com = new AccessComAutomation(_options);
        var view = includeValues ? AccessComAutomation.ResolveReportView("report") : AccessComAutomation.ResolveReportView("design");
        var closeSave = AccessComAutomation.ResolveCloseSave(false);

        var controls = await com.RunInAccessAsync(dbPath, app =>
        {
            var controlInfos = new List<object>();
            var doCmd = app.DoCmd;
            doCmd.OpenReport(reportName, view);

            // Held as `object`, not `dynamic` — see the note on AccessComAutomation's reflection helpers.
            object? report = null;
            object? controlsCollection = null;

            try
            {
                report = app.Reports[reportName];
                controlsCollection = AccessComAutomation.GetProperty(report!, "Controls")!;

                var count = AccessComAutomation.GetCount(controlsCollection);
                for (var i = 0; i < count; i++)
                {
                    var control = AccessComAutomation.GetItem(controlsCollection, i)!;
                    try
                    {
                        var name = Convert.ToString(AccessComAutomation.GetProperty(control, "Name")) ?? string.Empty;
                        var controlType = Convert.ToInt32(AccessComAutomation.GetProperty(control, "ControlType"), CultureInfo.InvariantCulture);
                        var controlSource = Convert.ToString(AccessComAutomation.TryGetProperty(control, "ControlSource"));
                        var value = includeValues ? AccessComAutomation.TryGetProperty(control, "Value") : null;

                        controlInfos.Add(new
                        {
                            name,
                            controlType,
                            controlTypeName = AccessComAutomation.GetControlTypeName(controlType),
                            controlSource,
                            value
                        });
                    }
                    finally
                    {
                        AccessComAutomation.ReleaseComObject(control);
                    }
                }
            }
            finally
            {
                try
                {
                    doCmd.Close(AccessComAutomation.ReportObjectType, reportName, closeSave);
                }
                catch
                {
                }

                AccessComAutomation.ReleaseComObject(controlsCollection);
                AccessComAutomation.ReleaseComObject(report);
                AccessComAutomation.ReleaseComObject(doCmd);
            }

            return (IReadOnlyList<object>)controlInfos;
        }, cancellationToken);

        return JsonSerializer.Serialize(new { databasePath = dbPath, reportName, controls }, JsonOptions);
    }

    [McpServerTool(Name = "access_add_report_control", Title = "Access: Add Report Control", ReadOnly = false, Idempotent = false)]
    [Description("Adds a control to an Access report via COM automation.")]
    public async Task<string> AddReportControlAsync(
        [Description("Report name.")] string reportName,
        [Description("Control type (e.g., textbox, combobox, checkbox).")]
        string controlType,
        [Description("Control name to assign.")] string controlName,
        [Description("Section (detail, header, footer, pageheader, pagefooter, grouplevel1header, grouplevel1footer, grouplevel2header, grouplevel2footer).")]
        string section = "detail",
        [Description("Left position (twips).")]
        double? left = null,
        [Description("Top position (twips).")]
        double? top = null,
        [Description("Width (twips).")]
        double? width = null,
        [Description("Height (twips).")]
        double? height = null,
        [Description("Path to .mdb/.accdb. If omitted, uses DefaultDatabasePath.")] string? databasePath = null,
        CancellationToken cancellationToken = default)
    {
        if (_options.ReadOnly)
        {
            throw new InvalidOperationException("Server is configured as read-only.");
        }

        ArgumentException.ThrowIfNullOrWhiteSpace(reportName);
        ArgumentException.ThrowIfNullOrWhiteSpace(controlType);
        ArgumentException.ThrowIfNullOrWhiteSpace(controlName);

        var dbPath = AccessDbPathPolicy.NormalizeAndValidate(databasePath, _options);
        var com = new AccessComAutomation(_options);
        var controlTypeValue = AccessComAutomation.ResolveControlType(controlType);
        var sectionValue = AccessComAutomation.ResolveSection(section);
        var closeSave = AccessComAutomation.ResolveCloseSave(true);

        await com.RunInAccessAsync(dbPath, app =>
        {
            var doCmd = app.DoCmd;
            doCmd.OpenReport(reportName, AccessComAutomation.ResolveReportView("design"));

            dynamic? created = null;

            try
            {
                created = app.CreateReportControl(
                    reportName,
                    controlTypeValue,
                    sectionValue,
                    Type.Missing,
                    Type.Missing,
                    left ?? Type.Missing,
                    top ?? Type.Missing,
                    width ?? Type.Missing,
                    height ?? Type.Missing);

                created.Name = controlName;
            }
            finally
            {
                try
                {
                    doCmd.Close(AccessComAutomation.ReportObjectType, reportName, closeSave);
                }
                catch
                {
                }

                AccessComAutomation.ReleaseComObject(created);
                AccessComAutomation.ReleaseComObject(doCmd);
            }

            return true;
        }, cancellationToken);

        return JsonSerializer.Serialize(new { databasePath = dbPath, reportName, controlName, controlType = controlTypeValue, section = sectionValue }, JsonOptions);
    }

    /// <summary>
    /// Removes a control from an Access report via COM automation.
    /// </summary>
    [McpServerTool(Name = "access_remove_report_control", Title = "Access: Remove Report Control", ReadOnly = false, Idempotent = false)]
    [Description("Removes a control from an Access report via COM automation.")]
    public async Task<string> RemoveReportControlAsync(
        [Description("Report name.")] string reportName,
        [Description("Control name.")] string controlName,
        [Description("Path to .mdb/.accdb. If omitted, uses DefaultDatabasePath.")] string? databasePath = null,
        CancellationToken cancellationToken = default)
    {
        if (_options.ReadOnly)
        {
            throw new InvalidOperationException("Server is configured as read-only.");
        }

        ArgumentException.ThrowIfNullOrWhiteSpace(reportName);
        ArgumentException.ThrowIfNullOrWhiteSpace(controlName);

        var dbPath = AccessDbPathPolicy.NormalizeAndValidate(databasePath, _options);
        var com = new AccessComAutomation(_options);
        var closeSave = AccessComAutomation.ResolveCloseSave(true);

        await com.RunInAccessAsync(dbPath, app =>
        {
            var doCmd = app.DoCmd;
            doCmd.OpenReport(reportName, AccessComAutomation.ResolveReportView("design"));

            try
            {
                app.DeleteControl(reportName, controlName);
            }
            finally
            {
                try
                {
                    doCmd.Close(AccessComAutomation.ReportObjectType, reportName, closeSave);
                }
                catch
                {
                }

                AccessComAutomation.ReleaseComObject(doCmd);
            }

            return true;
        }, cancellationToken);

        return JsonSerializer.Serialize(new { databasePath = dbPath, reportName, controlName }, JsonOptions);
    }

    [McpServerTool(Name = "access_run_macro", Title = "Access: Run Macro", ReadOnly = false, Idempotent = false)]
    [Description("Runs an Access macro by name via COM automation.")]
    public async Task<string> RunMacroAsync(
        [Description("Macro name.")] string macroName,
        [Description("Path to .mdb/.accdb. If omitted, uses DefaultDatabasePath.")] string? databasePath = null,
        CancellationToken cancellationToken = default)
    {
        if (_options.ReadOnly)
        {
            throw new InvalidOperationException("Server is configured as read-only.");
        }

        ArgumentException.ThrowIfNullOrWhiteSpace(macroName);

        var dbPath = AccessDbPathPolicy.NormalizeAndValidate(databasePath, _options);
        var com = new AccessComAutomation(_options);

        await com.RunInAccessAsync(dbPath, app =>
        {
            var doCmd = app.DoCmd;
            try
            {
                doCmd.RunMacro(macroName);
            }
            finally
            {
                AccessComAutomation.ReleaseComObject(doCmd);
            }

            return true;
        }, cancellationToken);

        return JsonSerializer.Serialize(new { databasePath = dbPath, macroName, ran = true }, JsonOptions);
    }

    [McpServerTool(Name = "access_do_cmd", Title = "Access: DoCmd", ReadOnly = true, Idempotent = false)]
    [Description("Runs a limited set of DoCmd actions via COM automation.")]
    public async Task<string> DoCmdAsync(
        [Description("Action: open_form, close_form, open_report, close_report.")] string action,
        [Description("Object name.")] string objectName,
        [Description("View for open actions (normal, design, report, preview, layout).")]
        string view = "normal",
        [Description("Path to .mdb/.accdb. If omitted, uses DefaultDatabasePath.")] string? databasePath = null,
        CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(action);
        ArgumentException.ThrowIfNullOrWhiteSpace(objectName);

        var normalizedAction = action.Trim().ToLowerInvariant();
        var dbPath = AccessDbPathPolicy.NormalizeAndValidate(databasePath, _options);
        var com = new AccessComAutomation(_options);

        await com.RunInAccessAsync(dbPath, app =>
        {
            var doCmd = app.DoCmd;
            try
            {
                switch (normalizedAction)
                {
                    case "open_form":
                        doCmd.OpenForm(objectName, AccessComAutomation.ResolveFormView(view.Equals("design", StringComparison.OrdinalIgnoreCase)));
                        break;
                    case "close_form":
                        doCmd.Close(AccessComAutomation.FormObjectType, objectName, AccessComAutomation.ResolveCloseSave(false));
                        break;
                    case "open_report":
                        doCmd.OpenReport(objectName, AccessComAutomation.ResolveReportView(view));
                        break;
                    case "close_report":
                        doCmd.Close(AccessComAutomation.ReportObjectType, objectName, AccessComAutomation.ResolveCloseSave(false));
                        break;
                    default:
                        throw new ArgumentException("Unsupported DoCmd action.", nameof(action));
                }
            }
            finally
            {
                AccessComAutomation.ReleaseComObject(doCmd);
            }

            return true;
        }, cancellationToken);

        return JsonSerializer.Serialize(new { databasePath = dbPath, action = normalizedAction, objectName, view }, JsonOptions);
    }

    [McpServerTool(Name = "access_execute_saved_query", Title = "Access: Execute Saved Query", ReadOnly = false, Idempotent = false)]
    [Description("Executes a saved Access query by name via COM automation. Returns rows if the query returns records.")]
    public async Task<string> ExecuteSavedQueryAsync(
        [Description("Saved query name.")] string queryName,
        [Description("Optional parameters as a JSON object.")] JsonElement? parameters = null,
        [Description("Path to .mdb/.accdb. If omitted, uses DefaultDatabasePath.")] string? databasePath = null,
        CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(queryName);

        var dbPath = AccessDbPathPolicy.NormalizeAndValidate(databasePath, _options);
        var args = ParseParameters(parameters);
        var com = new AccessComAutomation(_options);

        var result = await com.RunInAccessAsync<object>(dbPath, app =>
        {
            dynamic? db = null;
            dynamic? queryDef = null;
            dynamic? recordset = null;
            try
            {
                db = app.CurrentDb();
                queryDef = db.QueryDefs[queryName];

                if (args is not null && args.Count > 0)
                {
                    dynamic? paramsCollection = null;
                    try
                    {
                        paramsCollection = queryDef.Parameters;
                        foreach (var (name, value) in args)
                        {
                            try
                            {
                                paramsCollection[name].Value = value ?? DBNull.Value;
                            }
                            catch
                            {
                            }
                        }
                    }
                    finally
                    {
                        AccessComAutomation.ReleaseComObject(paramsCollection);
                    }
                }

                var returnsRecords = false;
                try
                {
                    returnsRecords = Convert.ToBoolean(queryDef.ReturnsRecords, CultureInfo.InvariantCulture);
                }
                catch
                {
                }

                if (returnsRecords)
                {
                    recordset = queryDef.OpenRecordset();
                    var columns = new List<string>();
                    var rows = new List<Dictionary<string, object?>>();

                    dynamic? fields = null;
                    try
                    {
                        fields = recordset.Fields;
                        var fieldCount = (int)fields.Count;
                        for (var i = 0; i < fieldCount; i++)
                        {
                            var field = fields.Item(i);
                            try
                            {
                                columns.Add(Convert.ToString(field.Name) ?? string.Empty);
                            }
                            finally
                            {
                                AccessComAutomation.ReleaseComObject(field);
                            }
                        }
                    }
                    finally
                    {
                        AccessComAutomation.ReleaseComObject(fields);
                    }

                    while (!recordset.EOF)
                    {
                        var row = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
                        for (var i = 0; i < columns.Count; i++)
                        {
                            object? value = null;
                            try
                            {
                                value = recordset.Fields[i].Value;
                            }
                            catch
                            {
                            }

                            if (value is DBNull)
                            {
                                value = null;
                            }

                            row[columns[i]] = value;
                        }

                        rows.Add(row);
                        recordset.MoveNext();
                    }

                    return new { returnsRecords = true, columns, rows };
                }

                if (_options.ReadOnly)
                {
                    throw new InvalidOperationException("Server is configured as read-only.");
                }

                queryDef.Execute();
                return new { returnsRecords = false, affected = (int?)null };
            }
            finally
            {
                try
                {
                    if (recordset is not null)
                    {
                        recordset.Close();
                    }
                }
                catch
                {
                }

                AccessComAutomation.ReleaseComObject(recordset);
                AccessComAutomation.ReleaseComObject(queryDef);
                AccessComAutomation.ReleaseComObject(db);
            }
        }, cancellationToken);

        return JsonSerializer.Serialize(new { databasePath = dbPath, queryName, result }, JsonOptions);
    }

    [McpServerTool(Name = "access_get_query_sql", Title = "Access: Get Query SQL", ReadOnly = true, Idempotent = true)]
    [Description("Gets a saved query's SQL text and type via COM automation, without executing it. Safe to call on action queries (append/update/delete/make-table).")]
    public async Task<string> GetQuerySqlAsync(
        [Description("Saved query name.")] string queryName,
        [Description("Path to .mdb/.accdb. If omitted, uses DefaultDatabasePath.")] string? databasePath = null,
        CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(queryName);

        var dbPath = AccessDbPathPolicy.NormalizeAndValidate(databasePath, _options);
        var com = new AccessComAutomation(_options);

        var result = await com.RunInAccessAsync(dbPath, app =>
        {
            dynamic? db = null;
            dynamic? queryDef = null;
            try
            {
                db = app.CurrentDb();
                queryDef = db.QueryDefs[queryName];

                var sql = Convert.ToString(queryDef.SQL) ?? string.Empty;

                var typeCode = 0;
                try
                {
                    typeCode = Convert.ToInt32(queryDef.Type, CultureInfo.InvariantCulture);
                }
                catch
                {
                }

                var returnsRecords = false;
                try
                {
                    returnsRecords = Convert.ToBoolean(queryDef.ReturnsRecords, CultureInfo.InvariantCulture);
                }
                catch
                {
                }

                return new { sql, queryType = DescribeQueryType(typeCode), returnsRecords };
            }
            finally
            {
                AccessComAutomation.ReleaseComObject(queryDef);
                AccessComAutomation.ReleaseComObject(db);
            }
        }, cancellationToken);

        return JsonSerializer.Serialize(new { databasePath = dbPath, queryName, result.sql, result.queryType, result.returnsRecords }, JsonOptions);
    }

    /// <summary>
    /// Maps a DAO QueryDef.Type constant (dbQSelect, dbQAppend, ...) to a readable name.
    /// </summary>
    private static string DescribeQueryType(int typeCode) => typeCode switch
    {
        0 => "Select",
        16 => "Crosstab",
        32 => "Delete",
        48 => "Update",
        64 => "Append",
        80 => "MakeTable",
        96 => "DDL",
        112 => "SQLPassThrough",
        128 => "Union",
        144 => "SQLPassThroughBulk",
        160 => "Compound",
        _ => $"Unknown ({typeCode})"
    };

    /// <summary>
    /// Resolves a VBA code module by name, falling back to the VBE object model.
    /// </summary>
    /// <remarks>
    /// Application.Modules only exposes standalone standard and class modules. A form's or report's
    /// code-behind (Form_Xxx / Report_Xxx) is not in that collection under any name — Access raises
    /// "can't find the module" even while the form is open in design view — so those resolve through
    /// the VBE's VBComponents instead, which needs Trust Center "Trust access to the VBA project
    /// object model". Both objects expose the CountOfLines/Lines/Proc* members the callers use.
    /// </remarks>
    private static object ResolveCodeModule(dynamic app, string moduleName)
    {
        try
        {
            object module = app.Modules[moduleName];
            return module;
        }
        catch (Exception modulesEx)
        {
            try
            {
                // The intermediate project/VBComponents/VBComponent wrappers are deliberately not
                // released: the CodeModule handed back to the caller is a child of them, and
                // FinalReleaseComObject on a parent invalidates it. RunInAccessAsync quits the Access
                // instance when the callback returns, which tears all of them down anyway.
                object vbe = app.VBE;
                object project = GetVbeProject(vbe, null);
                object components = AccessComAutomation.GetProperty(project, "VBComponents")!;
                object component = AccessComAutomation.GetItem(components, moduleName)!;
                return AccessComAutomation.GetProperty(component, "CodeModule")!;
            }
            catch (Exception vbeEx)
            {
                throw new InvalidOperationException(
                    $"Module '{moduleName}' is not in Application.Modules ({modulesEx.Message.Trim()}) and could not " +
                    $"be read via the VBE object model ({vbeEx.Message.Trim()}). Form and report code-behind is only " +
                    "reachable through the VBE, which requires Trust Center \"Trust access to the VBA project object model\".",
                    vbeEx);
            }
        }
    }

    private static dynamic GetVbeProject(dynamic vbe, string? projectName)
    {
        dynamic? projects = null;
        dynamic? selected = null;

        try
        {
            projects = vbe.VBProjects;
            var count = (int)projects.Count;
            // VBIDE collections are 1-based, unlike Access's own AllForms/AllModules/Controls.
            for (var i = 1; i <= count; i++)
            {
                var project = projects.Item(i);
                var name = Convert.ToString(project.Name) ?? string.Empty;
                var matches = projectName is null
                    ? selected is null
                    : name.Equals(projectName, StringComparison.OrdinalIgnoreCase);

                if (matches)
                {
                    if (selected is not null)
                    {
                        AccessComAutomation.ReleaseComObject(selected);
                    }

                    selected = project;
                }
                else
                {
                    AccessComAutomation.ReleaseComObject(project);
                }
            }

            if (selected is null)
            {
                throw new ArgumentException("VBA project not found.", nameof(projectName));
            }

            return selected;
        }
        finally
        {
            AccessComAutomation.ReleaseComObject(projects);
        }
    }

    private static string GetVbeComponentTypeName(int componentType)
    {
        return componentType switch
        {
            1 => "StandardModule",
            2 => "ClassModule",
            3 => "FormModule",
            100 => "DocumentModule",
            _ => componentType.ToString(CultureInfo.InvariantCulture)
        };
    }

    /// <summary>
    /// Converts a JSON object of parameter values into a dictionary for OLE DB commands.
    /// </summary>
    private static IReadOnlyDictionary<string, object?>? ParseParameters(JsonElement? parameters)
    {
        if (parameters is null || parameters.Value.ValueKind is JsonValueKind.Undefined or JsonValueKind.Null)
        {
            return null;
        }

        if (parameters.Value.ValueKind != JsonValueKind.Object)
        {
            throw new ArgumentException("Parameters must be a JSON object.", nameof(parameters));
        }

        var dict = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
        foreach (var prop in parameters.Value.EnumerateObject())
        {
            dict[prop.Name] = prop.Value.ValueKind switch
            {
                JsonValueKind.String => prop.Value.GetString(),
                JsonValueKind.Number => prop.Value.TryGetInt64(out var l) ? l : prop.Value.GetDouble(),
                JsonValueKind.True => true,
                JsonValueKind.False => false,
                JsonValueKind.Null => null,
                _ => prop.Value.GetRawText()
            };
        }

        return dict;
    }

    /// <summary>
    /// Converts a JSON object into a string-keyed dictionary for COM property setters.
    /// </summary>
    private static IReadOnlyDictionary<string, object?>? ParseJsonObject(JsonElement? properties)
    {
        if (properties is null || properties.Value.ValueKind is JsonValueKind.Undefined or JsonValueKind.Null)
        {
            return null;
        }

        if (properties.Value.ValueKind != JsonValueKind.Object)
        {
            throw new ArgumentException("Properties must be a JSON object.", nameof(properties));
        }

        var dict = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
        foreach (var prop in properties.Value.EnumerateObject())
        {
            dict[prop.Name] = ParseJsonValue(prop.Value);
        }

        return dict;
    }

    /// <summary>
    /// Converts a JSON element into a .NET value usable by Access COM automation.
    /// </summary>
    private static object? ParseJsonValue(JsonElement? value)
    {
        if (value is null || value.Value.ValueKind is JsonValueKind.Undefined or JsonValueKind.Null)
        {
            return null;
        }

        return value.Value.ValueKind switch
        {
            JsonValueKind.String => value.Value.GetString(),
            JsonValueKind.Number => value.Value.TryGetInt64(out var l) ? l : value.Value.GetDouble(),
            JsonValueKind.True => true,
            JsonValueKind.False => false,
            JsonValueKind.Null => null,
            _ => value.Value.GetRawText()
        };
    }

    private static readonly JsonSerializerOptions JsonOptions = new(JsonSerializerDefaults.Web)
    {
        WriteIndented = true
    };

    /// <summary>Like <see cref="JsonOptions"/>, but drops null properties - used to omit the optional
    /// unchanged-modules list from the module diff output unless it was requested.</summary>
    private static readonly JsonSerializerOptions OmitNullsJsonOptions = new(JsonSerializerDefaults.Web)
    {
        WriteIndented = true,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
    };

    /// <summary>
    /// Serializes rows into a simple CSV format for export.
    /// </summary>
    private static string SerializeCsv(IReadOnlyList<string> columns, IReadOnlyList<IReadOnlyDictionary<string, object?>> rows)
    {
        var builder = new StringBuilder();
        AppendCsvRow(builder, columns.Select(EscapeCsv));

        foreach (var row in rows)
        {
            var values = new string[columns.Count];
            for (var i = 0; i < columns.Count; i++)
            {
                row.TryGetValue(columns[i], out var value);
                values[i] = EscapeCsv(FormatValue(value));
            }

            AppendCsvRow(builder, values);
        }

        return builder.ToString();
    }

    /// <summary>
    /// Appends a single CSV row to the output builder.
    /// </summary>
    private static void AppendCsvRow(StringBuilder builder, IEnumerable<string> values)
    {
        var first = true;
        foreach (var value in values)
        {
            if (!first)
            {
                builder.Append(',');
            }

            builder.Append(value);
            first = false;
        }

        builder.AppendLine();
    }

    /// <summary>
    /// Escapes CSV values that contain delimiters or quotes.
    /// </summary>
    private static string EscapeCsv(string value)
    {
        var needsQuotes = value.Contains(',') || value.Contains('"') || value.Contains('\n') || value.Contains('\r');
        if (!needsQuotes)
        {
            return value;
        }

        return $"\"{value.Replace("\"", "\"\"")}\"";
    }

    /// <summary>
    /// Formats values for CSV output using invariant culture.
    /// </summary>
    private static string FormatValue(object? value)
    {
        return value switch
        {
            null => string.Empty,
            IFormattable formattable => formattable.ToString(null, CultureInfo.InvariantCulture) ?? string.Empty,
            _ => value.ToString() ?? string.Empty
        };
    }
}
