using DigitalJeeves.MsAccessMCP.Core;

namespace DigitalJeeves.MsAccessMCP.Tests;

public sealed class VbaTestKitTests
{
    [Fact]
    public void GenerateRunner_UsesDirectCallsAndHandlesSkip()
    {
        var tests = new List<(string, VbaProcedureInfo)>
        {
            ("M", Test("Passes", 1)),
            ("M", Test("Skipped", 2, skipped: true)),
            ("M", Test("Expected", 3, kind: "Function", expectedFailure: true))
        };

        var runner = VbaTestKit.GenerateRunner(tests);

        Assert.Contains("Call Passes", runner);
        Assert.Contains("Call Expected", runner);
        Assert.DoesNotContain("Call Skipped", runner);
        Assert.Contains("\"skip\"", runner);
        Assert.Contains("\"xfail\"", runner);
        Assert.DoesNotContain("Application.Run ", runner);
        Assert.Contains("vbNullChar", runner);
    }

    [Fact]
    public void GenerateRunner_RejectsDuplicateNames()
    {
        var tests = new List<(string, VbaProcedureInfo)>
        {
            ("M1", Test("Same", 1)),
            ("M2", Test("Same", 1))
        };

        var exception = Assert.Throws<InvalidOperationException>(() => VbaTestKit.GenerateRunner(tests));

        Assert.Contains("M1", exception.Message);
        Assert.Contains("M2", exception.Message);
    }

    [Fact]
    public void GenerateRunner_DuplicateCheckIgnoresCase()
    {
        // VBA names are case-insensitive, so these collide when the runner calls them.
        var tests = new List<(string, VbaProcedureInfo)>
        {
            ("M1", Test("Same", 1)),
            ("M2", Test("SAME", 1))
        };

        Assert.Throws<InvalidOperationException>(() => VbaTestKit.GenerateRunner(tests));
    }

    [Fact]
    public void GenerateRunner_ReportsPropertyTestsAsSkippedInsteadOfCallingThem()
    {
        var runner = VbaTestKit.GenerateRunner([("M", Test("Value", 1, kind: "Property Get"))]);

        Assert.DoesNotContain("Call Value", runner);
        Assert.Contains("Unsupported test procedure kind: Property Get", runner);
    }

    [Fact]
    public void GenerateRunner_EmitsTheMarkerTheNdjsonParserReads()
    {
        var runner = VbaTestKit.GenerateRunner([("M", Test("Passes", 1))]);

        Assert.Contains($"Public Function {VbaTestKit.RunnerEntryPoint}()", runner);
        Assert.Contains("\"@@test\" & TestSeparator & outcome & TestSeparator & testName", runner);
    }

    [Fact]
    public void GenerateRunner_WithNoTestsStillProducesARunner()
    {
        var runner = VbaTestKit.GenerateRunner([]);

        Assert.False(string.IsNullOrWhiteSpace(runner));
        Assert.DoesNotContain("Call ", runner);
    }

    [Fact]
    public void Kit_RendersResultPath()
    {
        var path = Path.Combine(Path.GetTempPath(), "result.ndjson");
        var rendered = VbaTestKit.Render(path);

        Assert.Contains(path.Replace("\"", "\"\""), rendered);
        Assert.DoesNotContain(VbaTestKit.ResultPathToken, rendered);
        Assert.Equal(64, VbaTestKit.Sha256.Length);
    }

    [Fact]
    public void Kit_RenderMakesRelativePathsAbsolute()
    {
        var rendered = VbaTestKit.Render("result.ndjson");

        Assert.Contains(Path.GetFullPath("result.ndjson"), rendered);
    }

    [Fact]
    public void Kit_RenderRefusesAnEmptyPath()
    {
        Assert.ThrowsAny<ArgumentException>(() => VbaTestKit.Render(" "));
    }

    [Theory]
    [InlineData("plain", "\"plain\"")]
    [InlineData("say \"hi\"", "\"say \"\"hi\"\"\"")]
    [InlineData("", "\"\"")]
    public void VbaString_QuotesAndDoublesEmbeddedQuotes(string value, string expected)
    {
        Assert.Equal(expected, VbaTestKit.VbaString(value));
    }

    private static VbaProcedureInfo Test(string name, int line, string kind = "Sub", bool skipped = false, bool expectedFailure = false)
        => new(name, kind, line, IsPublic: true, IsTest: true, skipped, expectedFailure);
}
