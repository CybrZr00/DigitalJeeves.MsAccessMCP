using System.Text.Json;
using DigitalJeeves.MsAccessMCP.Core;

namespace DigitalJeeves.MsAccessMCP.Tests;

/// <summary>
/// The OLE DB and DAO tools against a real, freshly created .accdb. Needs the ACE provider at the test
/// process's bitness; skipped otherwise. Access itself is never started.
/// </summary>
public sealed class AccessOleDbIntegrationTests : IDisposable
{
    private readonly TempDirectory _dir = new();

    public void Dispose() => _dir.Dispose();

    [AceFact]
    public async Task ExecuteThenQueryRoundTripsRows()
    {
        var (tools, db) = await CreateSampleAsync();

        using var result = Parse(await tools.QueryAsync("SELECT Name, Qty FROM Products ORDER BY Qty", db));
        var payload = result.RootElement.GetProperty("result");

        Assert.Equal(["Name", "Qty"], payload.GetProperty("columns").EnumerateArray().Select(c => c.GetString()));
        var rows = payload.GetProperty("rows").EnumerateArray().ToArray();
        Assert.Equal(3, rows.Length);
        Assert.Equal("Bolt", rows[0].GetProperty("Name").GetString());
        Assert.Equal(1, rows[0].GetProperty("Qty").GetInt32());
        Assert.Equal(JsonValueKind.Null, rows[2].GetProperty("Name").ValueKind);
    }

    [AceFact]
    public async Task ExecuteReportsAffectedRows()
    {
        var (tools, db) = await CreateSampleAsync();

        using var result = Parse(await tools.ExecuteAsync("UPDATE Products SET Qty = Qty + 1 WHERE Qty < 10", db));

        Assert.Equal(2, result.RootElement.GetProperty("affected").GetInt32());
    }

    [AceFact]
    public async Task ParametersArePositionalAndTheirNamesAreIgnored()
    {
        var (tools, db) = await CreateSampleAsync();

        // The keys deliberately don't match anything; only their order does.
        var parameters = Json("""{"zzz":"Widget","aaa":5}""");
        using var result = Parse(await tools.ScalarAsync("SELECT COUNT(*) FROM Products WHERE Name = ? AND Qty = ?", db, parameters));

        Assert.Equal(1, result.RootElement.GetProperty("value").GetInt32());
    }

    [AceFact]
    public async Task NullParameterIsSentAsDbNull()
    {
        var (tools, db) = await CreateSampleAsync();

        using var _ = Parse(await tools.ExecuteAsync("INSERT INTO Products (ID, Name, Qty) VALUES (?, ?, ?)", db, Json("""{"id":10,"name":null,"qty":0}""")));
        using var result = Parse(await tools.ScalarAsync("SELECT COUNT(*) FROM Products WHERE Name IS NULL", db));

        Assert.Equal(2, result.RootElement.GetProperty("value").GetInt32());
    }

    [AceFact]
    public async Task NonObjectParametersAreRejected()
    {
        var (tools, db) = await CreateSampleAsync();

        await Assert.ThrowsAsync<ArgumentException>(() => tools.QueryAsync("SELECT 1", db, Json("[1,2]")));
    }

    [AceFact]
    public async Task QueryPagedSlicesTheResult()
    {
        var (tools, db) = await CreateSampleAsync();

        using var page = Parse(await tools.QueryPagedAsync("SELECT ID FROM Products ORDER BY ID", pageNumber: 2, pageSize: 2, databasePath: db));
        using var beyond = Parse(await tools.QueryPagedAsync("SELECT ID FROM Products ORDER BY ID", pageNumber: 5, pageSize: 2, databasePath: db));

        Assert.Equal(3, page.RootElement.GetProperty("totalRows").GetInt32());
        Assert.Equal(3, Assert.Single(page.RootElement.GetProperty("rows").EnumerateArray()).GetProperty("ID").GetInt32());
        Assert.Equal(0, beyond.RootElement.GetProperty("rows").GetArrayLength());
    }

    [AceFact]
    public async Task QueryPagedValidatesPageArguments()
    {
        var (tools, db) = await CreateSampleAsync();

        await Assert.ThrowsAsync<ArgumentOutOfRangeException>(() => tools.QueryPagedAsync("SELECT 1", 0, 10, db));
        await Assert.ThrowsAsync<ArgumentOutOfRangeException>(() => tools.QueryPagedAsync("SELECT 1", 1, 0, db));
    }

    [AceFact]
    public async Task ExportTableAsCsvQuotesAwkwardValues()
    {
        var (tools, db) = await CreateSampleAsync();
        await tools.ExecuteAsync("UPDATE Products SET Name = 'Nut, \"hex\"' WHERE ID = 2", db);

        using var result = Parse(await tools.ExportTableAsync("Products", "CSV", databasePath: db));
        var lines = result.RootElement.GetProperty("csv").GetString()!.Split(Environment.NewLine, StringSplitOptions.RemoveEmptyEntries);

        Assert.Equal("ID,Name,Qty,Price", lines[0]);
        Assert.Contains(lines, l => l.StartsWith("2,\"Nut, \"\"hex\"\"\",1,", StringComparison.Ordinal));
        Assert.Equal("csv", result.RootElement.GetProperty("format").GetString());
    }

    [AceFact]
    public async Task ExportTableAsJsonHonoursMaxRows()
    {
        var (tools, db) = await CreateSampleAsync();

        using var result = Parse(await tools.ExportTableAsync("Products", "json", maxRows: 2, databasePath: db));

        Assert.Equal(2, result.RootElement.GetProperty("rows").GetArrayLength());
        await Assert.ThrowsAsync<ArgumentException>(() => tools.ExportTableAsync("Products", "xml", databasePath: db));
        await Assert.ThrowsAsync<ArgumentOutOfRangeException>(() => tools.ExportTableAsync("Products", "json", maxRows: 0, databasePath: db));
    }

    [AceFact]
    public async Task ListTablesExcludesSystemTables()
    {
        var (tools, db) = await CreateSampleAsync();

        using var result = Parse(await tools.ListTablesAsync(db));

        Assert.Equal(["Customers", "Products"], result.RootElement.GetProperty("tables").EnumerateArray().Select(t => t.GetString()));
    }

    [AceFact]
    public async Task GetSchemaListsColumnsInOrdinalOrder()
    {
        var (tools, db) = await CreateSampleAsync();

        using var result = Parse(await tools.GetSchemaAsync(db, "Products"));
        var table = Assert.Single(result.RootElement.GetProperty("schema").EnumerateArray());
        var columns = table.GetProperty("columns").EnumerateArray().ToArray();

        Assert.Equal("Products", table.GetProperty("table").GetString());
        Assert.Equal(["ID", "Name", "Qty", "Price"], columns.OrderBy(c => c.GetProperty("ordinal").GetInt32()).Select(c => c.GetProperty("name").GetString()));
        Assert.Equal(50, columns.Single(c => c.GetProperty("name").GetString() == "Name").GetProperty("characterMaximumLength").GetInt32());
    }

    [AceFact]
    public async Task ListIndexesReportsPrimaryKeys()
    {
        var (tools, db) = await CreateSampleAsync();

        using var indexes = Parse(await tools.ListIndexesAsync(db));
        var primaryKeyTables = indexes.RootElement.GetProperty("indexes").EnumerateArray()
            .Where(i => i.GetProperty("isPrimaryKey").GetBoolean())
            .Select(i => i.GetProperty("table").GetString())
            .Order();

        Assert.Equal(["Customers", "Products"], primaryKeyTables);
    }

    [AceFact]
    public async Task ListIndexesFiltersByTableName()
    {
        var (tools, db) = await CreateSampleAsync();

        using var indexes = Parse(await tools.ListIndexesAsync(db, "Products"));
        var tables = indexes.RootElement.GetProperty("indexes").EnumerateArray()
            .Select(i => i.GetProperty("table").GetString())
            .Distinct();

        Assert.Equal(["Products"], tables);
    }

    [AceFact]
    public async Task ListIndexesReportsAMultiColumnIndexOnce()
    {
        var (tools, db) = await CreateSampleAsync();
        await tools.ExecuteAsync("CREATE UNIQUE INDEX ixNameQty ON Products (Name, Qty)", db);

        using var indexes = Parse(await tools.ListIndexesAsync(db, "Products"));
        var index = Assert.Single(
            indexes.RootElement.GetProperty("indexes").EnumerateArray(),
            i => i.GetProperty("name").GetString() == "ixNameQty");

        Assert.True(index.GetProperty("isUnique").GetBoolean());
        Assert.False(index.GetProperty("isPrimaryKey").GetBoolean());
    }

    [AceFact]
    public async Task ListRelationshipsSeesForeignKeysAndFiltersByEitherTable()
    {
        var (tools, db) = await CreateSampleAsync();
        await tools.ExecuteAsync("ALTER TABLE Products ADD COLUMN CustomerID INTEGER", db);
        await tools.ExecuteAsync("ALTER TABLE Products ADD CONSTRAINT fkCustomer FOREIGN KEY (CustomerID) REFERENCES Customers (CustomerID)", db);

        using var all = Parse(await tools.ListRelationshipsAsync(db));
        using var byForeign = Parse(await tools.ListRelationshipsAsync(db, "products"));
        using var unrelated = Parse(await tools.ListRelationshipsAsync(db, "NoSuchTable"));

        var relationship = Assert.Single(all.RootElement.GetProperty("relationships").EnumerateArray());
        Assert.Equal("fkCustomer", relationship.GetProperty("name").GetString());
        Assert.Equal("Customers", relationship.GetProperty("primaryTable").GetString());
        Assert.Equal("CustomerID", relationship.GetProperty("primaryColumn").GetString());
        Assert.Equal("Products", relationship.GetProperty("foreignTable").GetString());
        Assert.Equal(1, byForeign.RootElement.GetProperty("relationships").GetArrayLength());
        Assert.Equal(0, unrelated.RootElement.GetProperty("relationships").GetArrayLength());
    }

    [AceFact]
    public async Task SavedQueriesShowUpAsViews()
    {
        var (tools, db) = await CreateSampleAsync();
        await tools.ExecuteAsync("CREATE VIEW qryBigOrders AS SELECT * FROM Products WHERE Qty > 2", db);

        using var views = Parse(await tools.ListViewsAsync(db));

        Assert.Contains("qryBigOrders", views.RootElement.GetProperty("views").EnumerateArray().Select(v => v.GetString()));
    }

    [AceFact]
    public async Task ReadOnlyServerStillQueriesButRefusesExecute()
    {
        var (_, db) = await CreateSampleAsync();
        var readOnly = new AccessTools(new AccessMcpOptions { ReadOnly = true });

        using var count = Parse(await readOnly.ScalarAsync("SELECT COUNT(*) FROM Products", db));

        Assert.Equal(3, count.RootElement.GetProperty("value").GetInt32());
        await Assert.ThrowsAsync<InvalidOperationException>(() => readOnly.ExecuteAsync("DELETE FROM Products", db));

        using var after = Parse(await readOnly.ScalarAsync("SELECT COUNT(*) FROM Products", db));
        Assert.Equal(3, after.RootElement.GetProperty("value").GetInt32());
    }

    [AceFact]
    public async Task DefaultDatabasePathIsUsedWhenNoneIsPassed()
    {
        var (_, db) = await CreateSampleAsync();
        var tools = new AccessTools(new AccessMcpOptions { DefaultDatabasePath = db });

        using var result = Parse(await tools.ScalarAsync("SELECT COUNT(*) FROM Products"));

        Assert.Equal(db, result.RootElement.GetProperty("databasePath").GetString());
        Assert.Equal(3, result.RootElement.GetProperty("value").GetInt32());
    }

    [AceFact]
    public async Task BadSqlSurfacesTheProviderError()
    {
        var (tools, db) = await CreateSampleAsync();

        await Assert.ThrowsAnyAsync<System.Data.OleDb.OleDbException>(() => tools.QueryAsync("SELECT * FROM NoSuchTable", db));
    }

    [AceFact]
    public async Task GetTableDesignReadsFieldsAndIndexesThroughDao()
    {
        var (tools, db) = await CreateSampleAsync();

        using var result = Parse(await tools.GetTableDesignAsync("Products", db));
        var table = Assert.Single(result.RootElement.GetProperty("tables").EnumerateArray());

        Assert.Equal("Products", table.GetProperty("name").GetString());
        Assert.Equal(4, table.GetProperty("fields").GetArrayLength());
        Assert.NotEqual(0, table.GetProperty("indexes").GetArrayLength());
    }

    private async Task<(AccessTools Tools, string Database)> CreateSampleAsync()
    {
        var db = TestDatabase.Create(_dir.Combine($"{Guid.NewGuid():N}.accdb"));
        var tools = new AccessTools(new AccessMcpOptions());

        await tools.ExecuteAsync("CREATE TABLE Customers (CustomerID INTEGER PRIMARY KEY, CustomerName TEXT(50))", db);
        await tools.ExecuteAsync("CREATE TABLE Products (ID INTEGER PRIMARY KEY, Name TEXT(50), Qty INTEGER, Price CURRENCY)", db);
        await tools.ExecuteAsync("INSERT INTO Products (ID, Name, Qty, Price) VALUES (1, 'Widget', 5, 2.5)", db);
        await tools.ExecuteAsync("INSERT INTO Products (ID, Name, Qty, Price) VALUES (2, 'Bolt', 1, 0.25)", db);
        await tools.ExecuteAsync("INSERT INTO Products (ID, Name, Qty, Price) VALUES (3, NULL, 20, 10)", db);

        return (tools, db);
    }

    private static JsonDocument Parse(string json) => JsonDocument.Parse(json);

    private static JsonElement Json(string json) => JsonDocument.Parse(json).RootElement.Clone();
}
