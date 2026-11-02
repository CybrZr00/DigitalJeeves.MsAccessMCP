using DigitalJeeves.MsAccessMCP.Core;

namespace DigitalJeeves.MsAccessMCP.Tests;

public sealed class DatabaseDiffTests
{
    [Fact]
    public void Compare_ClassifiesOnlyInSourceAndOnlyInTarget()
    {
        var result = VbaModuleDiff.Compare(
            [Module("ModuleOnlyInSource", "Sub A()\r\nEnd Sub")],
            [Module("ModuleOnlyInTarget", "Sub B()\r\nEnd Sub")],
            contextLines: 3);

        Assert.Equal("ModuleOnlyInSource", Assert.Single(result.OnlyInSource).Name);
        Assert.Equal("ModuleOnlyInTarget", Assert.Single(result.OnlyInTarget).Name);
        Assert.Empty(result.Changed);
        Assert.Empty(result.Unchanged);
    }

    [Fact]
    public void Compare_IdenticalModulesAreUnchanged()
    {
        var code = "Sub Main()\r\n    Debug.Print \"hi\"\r\nEnd Sub";

        var result = VbaModuleDiff.Compare([Module("Module1", code)], [Module("Module1", code)], contextLines: 3);

        Assert.Single(result.Unchanged);
        Assert.Empty(result.Changed);
        Assert.Empty(result.OnlyInSource);
        Assert.Empty(result.OnlyInTarget);
    }

    [Fact]
    public void Compare_MatchesModuleNamesIgnoringCase()
    {
        var result = VbaModuleDiff.Compare(
            [Module("Module1", "Sub A()\r\nEnd Sub")],
            [Module("MODULE1", "Sub A()\r\nEnd Sub")],
            contextLines: 3);

        Assert.Single(result.Unchanged);
        Assert.Empty(result.OnlyInSource);
        Assert.Empty(result.OnlyInTarget);
    }

    [Fact]
    public void Compare_ChangedModuleIncludesUnifiedDiffWithHeaders()
    {
        var source = Module("Module1", "Sub Main()\r\n    x = 1\r\nEnd Sub");
        var target = Module("Module1", "Sub Main()\r\n    x = 2\r\nEnd Sub");

        var result = VbaModuleDiff.Compare([source], [target], contextLines: 3);

        var change = Assert.Single(result.Changed);
        Assert.Equal("Module1", change.Name);
        Assert.Equal(3, change.SourceLineCount);
        Assert.Equal(3, change.TargetLineCount);
        Assert.Contains("--- source/Module1", change.Diff);
        Assert.Contains("+++ target/Module1", change.Diff);
        Assert.Contains("-    x = 1", change.Diff);
        Assert.Contains("+    x = 2", change.Diff);
    }

    [Fact]
    public void Compare_UsesTheSuppliedLabels()
    {
        var diff = VbaModuleDiff.Compare(
            [Module("Module1", "a")],
            [Module("Module1", "b")],
            contextLines: 3,
            sourceLabel: "live",
            targetLabel: "export").Changed[0].Diff;

        Assert.StartsWith("--- live/Module1\n+++ export/Module1\n", diff);
    }

    [Fact]
    public void Compare_ProducesExactUnifiedDiff()
    {
        var source = Module("Module1", "Sub Main()\r\n    x = 1\r\nEnd Sub");
        var target = Module("Module1", "Sub Main()\r\n    x = 2\r\nEnd Sub");

        var diff = VbaModuleDiff.Compare([source], [target], contextLines: 3).Changed[0].Diff;

        Assert.Equal(
            "--- source/Module1\n" +
            "+++ target/Module1\n" +
            "@@ -1,3 +1,3 @@\n" +
            " Sub Main()\n" +
            "-    x = 1\n" +
            "+    x = 2\n" +
            " End Sub\n",
            diff);
    }

    [Fact]
    public void Compare_PureInsertionAndDeletionProduceOneSidedHunks()
    {
        var inserted = VbaModuleDiff.Compare(
            [Module("M", "a\r\nc")],
            [Module("M", "a\r\nb\r\nc")],
            contextLines: 0).Changed[0].Diff;
        var deleted = VbaModuleDiff.Compare(
            [Module("M", "a\r\nb\r\nc")],
            [Module("M", "a\r\nc")],
            contextLines: 0).Changed[0].Diff;

        Assert.Contains("+b\n", inserted);
        Assert.DoesNotContain("\n-", inserted);
        Assert.Contains("-b\n", deleted);
        Assert.DoesNotContain("\n+", deleted.Substring(deleted.IndexOf("@@", StringComparison.Ordinal)));
    }

    [Fact]
    public void Compare_DistantChangesProduceSeparateHunks()
    {
        var sourceLines = Enumerable.Range(1, 30).Select(i => $"line{i}").ToArray();
        var targetLines = (string[])sourceLines.Clone();
        targetLines[1] = "changed2";
        targetLines[27] = "changed28";

        var diff = VbaModuleDiff.Compare(
            [Module("M", string.Join("\r\n", sourceLines))],
            [Module("M", string.Join("\r\n", targetLines))],
            contextLines: 2).Changed[0].Diff;

        Assert.Equal(2, diff.Split('\n').Count(l => l.StartsWith("@@", StringComparison.Ordinal)));
    }

    [Fact]
    public void Compare_CaseOnlyDifferenceIsIgnoredByDefaultButReportedWhenCaseSensitive()
    {
        var source = Module("Module1", "Sub Main()\r\n    web = rstCompany!web\r\nEnd Sub");
        var target = Module("Module1", "Sub Main()\r\n    Web = rstCompany!Web\r\nEnd Sub");

        var ignoring = VbaModuleDiff.Compare([source], [target], contextLines: 3);
        Assert.Empty(ignoring.Changed);
        Assert.Single(ignoring.Unchanged);

        var strict = VbaModuleDiff.Compare([source], [target], contextLines: 3, ignoreCase: false);
        Assert.Single(strict.Changed);
        Assert.Empty(strict.Unchanged);
    }

    [Fact]
    public void Compare_IgnoreCaseHidesCaseOnlyLinesButKeepsRealChanges()
    {
        var source = Module("Module1", "Sub Main()\r\n    Web = 1\r\n    x = 1\r\nEnd Sub");
        var target = Module("Module1", "Sub Main()\r\n    web = 1\r\n    x = 2\r\nEnd Sub");

        var diff = VbaModuleDiff.Compare([source], [target], contextLines: 0, ignoreCase: true).Changed[0].Diff;

        Assert.Contains("-    x = 1", diff);
        Assert.Contains("+    x = 2", diff);
        Assert.DoesNotContain("Web", diff);
        Assert.DoesNotContain("web", diff);
    }

    [Fact]
    public void Compare_ContextLinesControlHunkSize()
    {
        var sourceLines = Enumerable.Range(1, 10).Select(i => $"line{i}").ToArray();
        var targetLines = (string[])sourceLines.Clone();
        targetLines[4] = "changed";

        var source = Module("Module1", string.Join("\r\n", sourceLines));
        var target = Module("Module1", string.Join("\r\n", targetLines));

        var withContext = VbaModuleDiff.Compare([source], [target], contextLines: 3).Changed[0].Diff;
        var withoutContext = VbaModuleDiff.Compare([source], [target], contextLines: 0).Changed[0].Diff;

        Assert.Contains("@@ -2,7 +2,7 @@", withContext);
        Assert.Contains("@@ -5 +5 @@", withoutContext);
        // Context 0 must not leak any surrounding equal lines into the hunk.
        Assert.DoesNotContain(" line4", withoutContext);
    }

    [Fact]
    public void Compare_NegativeContextIsRejected()
    {
        Assert.Throws<ArgumentOutOfRangeException>(
            () => VbaModuleDiff.Compare([Module("M", "a")], [Module("M", "b")], contextLines: -1));
    }

    [Fact]
    public void Compare_WhitespaceOnlyDifferenceIsIgnoredByDefaultButReportedWhenStrict()
    {
        var source = Module("Module1", "Sub Main()\r\n    x = 1\r\nEnd Sub");
        var target = Module("Module1", "Sub Main()\r\n\tx  =  1   \r\nEnd Sub");

        var ignoring = VbaModuleDiff.Compare([source], [target], contextLines: 3);
        Assert.Empty(ignoring.Changed);
        Assert.Single(ignoring.Unchanged);

        var strict = VbaModuleDiff.Compare([source], [target], contextLines: 3, ignoreWhitespace: false);
        Assert.Single(strict.Changed);
        Assert.Empty(strict.Unchanged);
    }

    [Fact]
    public void Compare_IgnoreWhitespaceHidesWhitespaceOnlyLinesButKeepsRealChanges()
    {
        var source = Module("Module1", "Sub Main()\r\n    x  =  1\r\n    y = 2\r\nEnd Sub");
        var target = Module("Module1", "Sub Main()\r\n\tx = 1\r\n    y = 3\r\nEnd Sub");

        var diff = VbaModuleDiff.Compare([source], [target], contextLines: 0).Changed[0].Diff;

        Assert.Contains("-    y = 2", diff);
        Assert.Contains("+    y = 3", diff);
        Assert.DoesNotContain("x", diff);
    }

    [Fact]
    public void Compare_SpacingAddedWhereThereWasNoneIsStillAChange()
    {
        var source = Module("Module1", "Sub Main()\r\n    x = 1\r\nEnd Sub");
        var target = Module("Module1", "Sub Main()\r\n    x=1\r\nEnd Sub");

        var result = VbaModuleDiff.Compare([source], [target], contextLines: 3);

        Assert.Single(result.Changed);
        Assert.Empty(result.Unchanged);
    }

    [Fact]
    public void Compare_LineEndingOnlyDifferenceIsUnchanged()
    {
        var source = Module("Module1", "Sub Main()\r\n    x = 1\r\nEnd Sub");
        var target = Module("Module1", "Sub Main()\n    x = 1\nEnd Sub");

        var result = VbaModuleDiff.Compare([source], [target], contextLines: 3);

        Assert.Empty(result.Changed);
        Assert.Single(result.Unchanged);
    }

    [Fact]
    public void Compare_IgnoreCaseAndWhitespaceCombine()
    {
        var source = Module("Module1", "Sub Main()\r\n    Web  =  1\r\nEnd Sub");
        var target = Module("Module1", "Sub Main()\r\n\tweb = 1\r\nEnd Sub");

        Assert.Empty(VbaModuleDiff.Compare([source], [target], contextLines: 3).Changed);
        Assert.Single(VbaModuleDiff.Compare([source], [target], contextLines: 3, ignoreCase: false).Changed);
        Assert.Single(VbaModuleDiff.Compare([source], [target], contextLines: 3, ignoreWhitespace: false).Changed);
    }

    [Fact]
    public void Compare_LargeUnrelatedModulesStillProduceADiff()
    {
        // Big enough to exercise the linear-space LCS path without hitting the cell cap.
        var source = string.Join("\r\n", Enumerable.Range(0, 2000).Select(i => $"src{i}"));
        var target = string.Join("\r\n", Enumerable.Range(0, 2000).Select(i => $"dst{i}"));

        var change = Assert.Single(VbaModuleDiff.Compare([Module("M", source)], [Module("M", target)], contextLines: 0).Changed);

        Assert.Equal(2000, change.Diff.Split('\n').Count(l => l.StartsWith('-') && !l.StartsWith("---", StringComparison.Ordinal)));
        Assert.Equal(2000, change.Diff.Split('\n').Count(l => l.StartsWith('+') && !l.StartsWith("+++", StringComparison.Ordinal)));
    }

    [Fact]
    public void QueryDiff_ClassifiesOnlyInSourceAndOnlyInTarget()
    {
        var result = QueryDiff.Compare(
            [Query("QueryOnlyInSource", "SELECT * FROM A")],
            [Query("QueryOnlyInTarget", "SELECT * FROM B")],
            contextLines: 3);

        Assert.Equal("QueryOnlyInSource", Assert.Single(result.OnlyInSource).Name);
        Assert.Equal("QueryOnlyInTarget", Assert.Single(result.OnlyInTarget).Name);
        Assert.Empty(result.Changed);
        Assert.Empty(result.Unchanged);
    }

    [Fact]
    public void QueryDiff_ChangedQueryIncludesUnifiedDiffWithHeaders()
    {
        var source = Query("Qry1", "SELECT A\r\nFROM T");
        var target = Query("Qry1", "SELECT B\r\nFROM T");

        var change = Assert.Single(QueryDiff.Compare([source], [target], contextLines: 3).Changed);

        Assert.Equal(42, change.QueryType);
        Assert.Equal(2, change.SourceLineCount);
        Assert.Equal(2, change.TargetLineCount);
        Assert.Contains("--- source/Qry1", change.Diff);
        Assert.Contains("+++ target/Qry1", change.Diff);
        Assert.Contains("-SELECT A", change.Diff);
        Assert.Contains("+SELECT B", change.Diff);
    }

    [Fact]
    public void QueryDiff_IgnoresCaseAndWhitespaceByDefault()
    {
        var source = Query("Qry1", "SELECT  [A]  \r\nFROM [T]");
        var target = Query("Qry1", "select [a]\r\nfrom [t]");

        var ignoring = QueryDiff.Compare([source], [target], contextLines: 3);
        Assert.Empty(ignoring.Changed);
        Assert.Single(ignoring.Unchanged);

        var caseStrict = QueryDiff.Compare([source], [target], contextLines: 3, ignoreCase: false);
        Assert.Single(caseStrict.Changed);

        var whitespaceStrict = QueryDiff.Compare([source], [target], contextLines: 3, ignoreWhitespace: false);
        Assert.Single(whitespaceStrict.Changed);
    }

    [Fact]
    public void CountLines_HandlesEmptyAndMultilineBodies()
    {
        Assert.Equal(0, VbaModuleDiff.CountLines(string.Empty));
        Assert.Equal(1, VbaModuleDiff.CountLines("Sub A()"));
        Assert.Equal(3, VbaModuleDiff.CountLines("a\r\nb\r\nc"));
    }

    private static VbaModuleSnapshot Module(string name, string code) => new(name, ComponentType: 1, code);

    private static QuerySnapshot Query(string name, string sql) => new(name, QueryType: 42, sql);
}
