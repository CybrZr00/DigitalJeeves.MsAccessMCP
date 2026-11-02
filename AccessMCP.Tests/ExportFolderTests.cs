using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using DigitalJeeves.MsAccessMCP.Core;

namespace DigitalJeeves.MsAccessMCP.Tests;

/// <summary>
/// The file-format and drift logic behind access_export_all_modules_to_folder,
/// access_check_export_drift and access_import_changed_modules_from_folder.
/// </summary>
public sealed class ExportFolderTests : IDisposable
{
    private readonly TempDirectory _dir = new();

    public void Dispose() => _dir.Dispose();

    [Fact]
    public void ExportManifest_OldJsonWithoutQueriesDeserializes()
    {
        const string json = """
            {
              "databasePath": "C:\\Data\\Sample.accdb",
              "projectName": "VBAProject",
              "exportedAtUtc": "2024-01-01T00:00:00+00:00",
              "modules": []
            }
            """;

        var manifest = JsonSerializer.Deserialize<AccessTools.ExportManifest>(
            json,
            new JsonSerializerOptions(JsonSerializerDefaults.Web));

        Assert.NotNull(manifest);
        Assert.Null(manifest!.Queries);
    }

    [Fact]
    public void QueryFileBody_RoundTripsThroughNormalization()
    {
        var sql = "SELECT A\nFROM T\r\nWHERE B = 1";
        var normalized = AccessTools.NormalizeQuerySql(sql);

        Assert.Equal(normalized, AccessTools.ReadExportedQueryBody(normalized + "\r\n"));
        Assert.DoesNotContain("\n", normalized.Replace("\r\n", string.Empty));
    }

    [Fact]
    public void NormalizeQuerySql_ConvertsLoneCarriageReturns()
    {
        Assert.Equal("a\r\nb\r\nc", AccessTools.NormalizeQuerySql("a\rb\nc"));
    }

    [Fact]
    public void ReadExportedQueryBody_DropsOnlyOneTrailingNewline()
    {
        // A query whose SQL itself ends in a blank line keeps it.
        Assert.Equal("SELECT 1\r\n", AccessTools.ReadExportedQueryBody("SELECT 1\n\n"));
        Assert.Equal("SELECT 1", AccessTools.ReadExportedQueryBody("SELECT 1"));
    }

    [Fact]
    public void QueryFileNames_DisambiguateSanitizedCollisions()
    {
        var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        var first = AccessTools.GetUniqueQueryFileName("A/B", used);
        var second = AccessTools.GetUniqueQueryFileName("A_B", used);

        Assert.NotEqual(first, second, StringComparer.OrdinalIgnoreCase);
        Assert.Contains("~", second);
    }

    [Fact]
    public void QueryFileNames_AreDeterministicAndLiveUnderQueries()
    {
        var name = AccessTools.GetUniqueQueryFileName("qry:Totals?", new HashSet<string>(StringComparer.OrdinalIgnoreCase));
        var again = AccessTools.GetUniqueQueryFileName("qry:Totals?", new HashSet<string>(StringComparer.OrdinalIgnoreCase));

        Assert.Equal(Path.Combine("Queries", "qry_Totals_.sql"), name);
        Assert.Equal(name, again);
    }

    [Fact]
    public void QueryFileNames_ThirdCollisionGetsANumericSuffix()
    {
        var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        var names = new[]
        {
            AccessTools.GetUniqueQueryFileName("Q", used),
            AccessTools.GetUniqueQueryFileName("Q", used),
            AccessTools.GetUniqueQueryFileName("Q", used)
        };

        Assert.Equal(3, names.Distinct(StringComparer.OrdinalIgnoreCase).Count());
        Assert.EndsWith("-1.sql", names[2]);
    }

    [Fact]
    public void StripExportHeaders_RemovesClassHeaderAndAttributes()
    {
        const string exported =
            "VERSION 1.0 CLASS\r\n" +
            "BEGIN\r\n" +
            "  MultiUse = -1  'True\r\n" +
            "END\r\n" +
            "Attribute VB_Name = \"clsThing\"\r\n" +
            "Attribute VB_GlobalNameSpace = False\r\n" +
            "Option Compare Database\r\n" +
            "Option Explicit\r\n" +
            "\r\n" +
            "Public Sub Go()\r\n" +
            "End Sub\r\n";

        Assert.Equal(
            "Option Compare Database\r\nOption Explicit\r\n\r\nPublic Sub Go()\r\nEnd Sub",
            AccessTools.StripExportHeaders(exported));
    }

    [Fact]
    public void StripExportHeaders_SkipsNestedBeginBlocksInFormExports()
    {
        const string exported =
            "VERSION 5.00\r\n" +
            "Begin Form\r\n" +
            "    Begin Section\r\n" +
            "    End\r\n" +
            "End\r\n" +
            "Attribute VB_Name = \"Form_Main\"\r\n" +
            "Private Sub Form_Load()\r\n" +
            "End Sub\r\n";

        // "End" closes a block only when it stands alone, and only the outermost one ends the header.
        var body = AccessTools.StripExportHeaders(exported);

        Assert.StartsWith("Private Sub Form_Load()", body);
    }

    [Fact]
    public void StripExportHeaders_KeepsAFinalBlankLineThatIsPartOfTheBody()
    {
        // The file's own terminating newline is removed exactly once.
        Assert.Equal("Sub A()\r\nEnd Sub\r\n", AccessTools.StripExportHeaders("Attribute VB_Name = \"M\"\nSub A()\nEnd Sub\n\n"));
    }

    [Fact]
    public void ClassifyDrift_SortsEveryModuleIntoTheRightBucket()
    {
        const string original = "Sub A()\r\nEnd Sub";
        const string edited = "Sub A()\r\n    x = 1\r\nEnd Sub";
        const string otherEdit = "Sub A()\r\n    x = 2\r\nEnd Sub";

        var manifest = Manifest(
            Entry("Unchanged", original),
            Entry("Drifted", original),
            Entry("EditedLocally", original),
            Entry("BothSame", original),
            Entry("Conflict", original),
            Entry("Removed", original),
            Entry("MissingFile", original));

        WriteModule("Unchanged", original);
        WriteModule("Drifted", original);
        WriteModule("EditedLocally", edited);
        WriteModule("BothSame", edited);
        WriteModule("Conflict", edited);
        WriteModule("Removed", original);

        var live = new List<AccessTools.LiveModule>
        {
            Live("Unchanged", original),
            Live("Drifted", edited),
            Live("EditedLocally", original),
            Live("BothSame", edited),
            Live("Conflict", otherEdit),
            Live("MissingFile", original),
            Live("BrandNew", original)
        };

        var report = AccessTools.ClassifyDrift(_dir.Path, manifest, live);

        Assert.Equal(1, report.UnchangedCount);
        Assert.Equal(["Drifted"], report.DriftedInAccess);
        Assert.Equal(["EditedLocally"], report.EditedLocally);
        Assert.Equal(["BothSame"], report.LocalMatchesLive);
        Assert.Equal(["Conflict"], report.Conflicts);
        Assert.Equal(["Removed"], report.RemovedFromAccess);
        Assert.Equal(["MissingFile"], report.MissingFiles);
        Assert.Equal(["BrandNew"], report.NewInAccess);
    }

    [Fact]
    public void ClassifyDrift_MissingFileThatAlsoDriftedIsReportedTwice()
    {
        var manifest = Manifest(Entry("M", "Sub A()\r\nEnd Sub"));
        var live = new List<AccessTools.LiveModule> { Live("M", "Sub B()\r\nEnd Sub") };

        var report = AccessTools.ClassifyDrift(_dir.Path, manifest, live);

        Assert.Equal(["M"], report.MissingFiles);
        Assert.Equal(["M"], report.DriftedInAccess);
    }

    [Fact]
    public void ClassifyDrift_ExportHeadersAndNameCaseDoNotCountAsChanges()
    {
        const string body = "Option Explicit\r\n\r\nSub A()\r\nEnd Sub";
        var manifest = Manifest(Entry("Module1", body));
        File.WriteAllText(_dir.Combine("Module1.bas"), "Attribute VB_Name = \"Module1\"\r\n" + body + "\r\n");

        var report = AccessTools.ClassifyDrift(_dir.Path, manifest, [Live("MODULE1", body)]);

        Assert.Equal(1, report.UnchangedCount);
        Assert.Empty(report.NewInAccess);
    }

    [Fact]
    public void ClassifyDrift_QueryEntriesAreReadAsSql()
    {
        const string sql = "SELECT *\r\nFROM T";
        var entry = new AccessTools.ExportManifestEntry("qryAll", 0, "qryAll.sql", "sql", true, Hash(sql));
        File.WriteAllText(_dir.Combine("qryAll.sql"), "SELECT *\nFROM T\n");

        var report = AccessTools.ClassifyDrift(_dir.Path, Manifest(entry), [Live("qryAll", sql)]);

        Assert.Equal(1, report.UnchangedCount);
    }

    private static AccessTools.ExportManifest Manifest(params AccessTools.ExportManifestEntry[] entries)
        => new(@"C:\Data\Sample.accdb", "VBAProject", DateTimeOffset.UnixEpoch, [.. entries]);

    private static AccessTools.ExportManifestEntry Entry(string name, string body)
        => new(name, 1, name + ".bas", "export", true, Hash(body));

    private static AccessTools.LiveModule Live(string name, string body) => new(name, 1, Hash(body));

    private void WriteModule(string name, string body)
        => File.WriteAllText(_dir.Combine(name + ".bas"), $"Attribute VB_Name = \"{name}\"\r\n{body}\r\n");

    // Matches AccessTools.Sha256Hex: the raw UTF-8 body, no normalization.
    private static string Hash(string text) => Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(text)));
}
