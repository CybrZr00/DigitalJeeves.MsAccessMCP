using DigitalJeeves.MsAccessMCP.Core;

namespace DigitalJeeves.MsAccessMCP.Tests;

public sealed class NdjsonResultsTests : IDisposable
{
    private readonly TempDirectory _dir = new();

    public void Dispose() => _dir.Dispose();

    [Fact]
    public void ParseTests_ParsesOutcomesAndIncompleteTests()
    {
        var path = Write(
            """{"event":"log","text":"@@test\u0000start\u0000A\u0000100\u0000"}""",
            """{"event":"log","text":"@@test\u0000pass\u0000A\u0000105\u0000"}""",
            """{"event":"log","text":"@@test\u0000start\u0000B\u0000200\u0000"}""",
            """{"event":"log","text":"@@test\u0000fail\u0000B\u0000210\u0000boom"}""",
            """{"event":"log","text":"@@test\u0000skip\u0000C\u00000\u0000"}""",
            """{"event":"log","text":"hello"}""",
            """{"event":"log","text":"@@test\u0000start\u0000D\u0000300\u0000"}""",
            """{"event":"log","text":"@@test\u0000pass\u0000D\u0000305\u0000"}trailing""");

        var summary = NdjsonResults.ParseTests(path, 100_000, markIncomplete: true);

        Assert.Equal(4, summary.Total);
        Assert.Equal(1, summary.Passed);
        Assert.Equal(2, summary.Failed);
        Assert.Equal(1, summary.Skipped);
        Assert.Contains(summary.Tests, t => t.Name == "A" && t.ElapsedMs == 5);
        Assert.Contains(summary.Tests, t => t.Name == "B" && t.Message == "boom");
        Assert.Contains(summary.Tests, t => t.Name == "D" && t.Outcome == "error");
        Assert.Contains("hello", summary.LogTail);
        Assert.False(summary.Ok);
    }

    [Fact]
    public void ParseTests_StartedButUnfinishedTestIsDroppedUnlessMarked()
    {
        var path = Write(
            """{"event":"log","text":"@@test\u0000start\u0000A\u0000100\u0000"}""");

        var unmarked = NdjsonResults.ParseTests(path, 100_000);
        var marked = NdjsonResults.ParseTests(path, 100_000, markIncomplete: true);

        Assert.Equal(0, unmarked.Total);
        Assert.Equal(1, marked.Total);
        Assert.False(marked.Ok);
    }

    [Fact]
    public void ParseTests_XfailIsNotAFailureButXpassIs()
    {
        var path = Write(
            """{"event":"log","text":"@@test\u0000xfail\u0000A\u00000\u0000expected"}""",
            """{"event":"log","text":"@@test\u0000pass\u0000B\u00000\u0000"}""");

        var summary = NdjsonResults.ParseTests(path, 100_000);
        Assert.True(summary.Ok);
        Assert.Equal(1, summary.XFailed);
        Assert.Equal(0, summary.Failed);

        var xpassPath = Write(
            """{"event":"log","text":"@@test\u0000xpass\u0000A\u00000\u0000"}""");

        var xpass = NdjsonResults.ParseTests(xpassPath, 100_000);
        Assert.False(xpass.Ok);
        Assert.Equal(1, xpass.XPassed);
        Assert.Equal(1, xpass.Failed);
    }

    [Fact]
    public void ParseTests_AllPassingRunIsOk()
    {
        var path = Write(
            """{"event":"log","text":"@@test\u0000start\u0000A\u0000100\u0000"}""",
            """{"event":"log","text":"@@test\u0000pass\u0000A\u0000101\u0000"}""",
            """{"event":"log","text":"@@test\u0000skip\u0000B\u00000\u0000"}""");

        var summary = NdjsonResults.ParseTests(path, 100_000, markIncomplete: true);

        Assert.True(summary.Ok);
        Assert.Null(summary.Error);
        Assert.Equal(["A", "B"], summary.Tests.Select(t => t.Name));
    }

    [Fact]
    public void ParseTests_ErrorEventFailsTheRunAndKeepsTheFirstError()
    {
        var path = Write(
            """{"event":"log","text":"@@test\u0000pass\u0000A\u00000\u0000"}""",
            """{"event":"error","source":"Runner","description":"first"}""",
            """{"event":"error","source":"Runner","description":"second"}""");

        var summary = NdjsonResults.ParseTests(path, 100_000);

        Assert.False(summary.Ok);
        Assert.Equal("Runner: first", summary.Error);
    }

    [Fact]
    public void ParseTests_AcceptsTheOlderFormatWithoutTickCounts()
    {
        var path = Write(
            """{"event":"log","text":"@@test\u0000fail\u0000A\u0000it broke"}""");

        var outcome = Assert.Single(NdjsonResults.ParseTests(path, 100_000).Tests);

        Assert.Equal("fail", outcome.Outcome);
        Assert.Equal("it broke", outcome.Message);
        Assert.Equal(0, outcome.ElapsedMs);
    }

    [Fact]
    public void ParseTests_LaterOutcomeReplacesEarlierForTheSameTest()
    {
        var path = Write(
            """{"event":"log","text":"@@test\u0000fail\u0000A\u00000\u0000first"}""",
            """{"event":"log","text":"@@test\u0000pass\u0000a\u00000\u0000"}""");

        var outcome = Assert.Single(NdjsonResults.ParseTests(path, 100_000).Tests);

        Assert.Equal("pass", outcome.Outcome);
    }

    [Fact]
    public void ParseTests_ValueEventsGoToTheLog()
    {
        var path = Write("""{"event":"value","type":"Long","text":"42"}""");

        var summary = NdjsonResults.ParseTests(path, 100_000);

        Assert.Equal(["[value] Long: 42"], summary.LogTail);
    }

    [Fact]
    public void ParseTests_MissingFileIsAnEmptyOkSummary()
    {
        var summary = NdjsonResults.ParseTests(_dir.Combine("missing.ndjson"), 100_000);

        Assert.True(summary.Ok);
        Assert.Equal(0, summary.Total);
        Assert.Empty(summary.LogTail);
    }

    [Fact]
    public void ParseLog_KeepsTheNewestLinesWithinTheCap()
    {
        var lines = Enumerable.Range(1, 20)
            .Select(i => $$"""{"event":"log","text":"line{{i:D2}}"}""")
            .ToArray();
        var path = Write(lines);

        // Each log line is 6 chars + 1 separator; 21 chars leaves room for the last three.
        var (log, _) = NdjsonResults.ParseLog(path, 21);

        Assert.Equal(["line18", "line19", "line20"], log);
    }

    [Fact]
    public void ParseLog_SkipsTestMarkersAndReportsTheLastError()
    {
        var path = Write(
            """{"event":"log","text":"@@test\u0000pass\u0000A\u00000\u0000"}""",
            """{"event":"log","text":"plain"}""",
            """{"event":"value","type":"String","text":"x"}""",
            """{"event":"error","source":"S","description":"first"}""",
            """{"event":"error","source":"S","description":"last"}""",
            """{"event":"log","text":"partial""");

        var (log, error) = NdjsonResults.ParseLog(path, 100_000);

        Assert.Equal(["plain", "[value] String: x"], log);
        Assert.Equal("S: last", error);
    }

    private string Write(params string[] lines)
    {
        var path = _dir.Combine($"{Guid.NewGuid():N}.ndjson");
        File.WriteAllLines(path, lines);
        return path;
    }
}
