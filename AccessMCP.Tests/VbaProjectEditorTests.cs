using DigitalJeeves.MsAccessMCP.Core;

namespace DigitalJeeves.MsAccessMCP.Tests;

public sealed class VbaProjectEditorTests
{
    [Fact]
    public void PrepareForWrite_ExpandsTabsToEditorTabStops()
    {
        var code = "Public Sub Probe()\r\n\tx = 1\r\n  \tIf x Then\r\nEnd Sub";

        var prepared = VbaProjectEditor.PrepareForWrite(code, string.Empty);

        var lines = prepared.Code.Split("\r\n");
        Assert.Equal("Public Sub Probe()", lines[0]);
        Assert.Equal("    x = 1", lines[1]);
        Assert.Equal("    If x Then", lines[2]);
        Assert.Equal("End Sub", lines[3]);
        Assert.Equal(2, prepared.ExpandedTabs);
        Assert.DoesNotContain('\t', prepared.Code);
    }

    [Fact]
    public void PrepareForWrite_TabMidLineAdvancesToTheNextStop()
    {
        var prepared = VbaProjectEditor.PrepareForWrite("x\t= 1\r\nab\tc", string.Empty);

        Assert.Equal("x   = 1\r\nab  c", prepared.Code);
    }

    [Fact]
    public void PrepareForWrite_TrimsTrailingWhitespaceButKeepsWhitespaceOnlyLines()
    {
        var code = "Public Sub Probe()\r\n    x = 1   \r\n    \r\nEnd Sub";

        var prepared = VbaProjectEditor.PrepareForWrite(code, string.Empty);

        Assert.Equal("Public Sub Probe()\r\n    x = 1\r\n    \r\nEnd Sub", prepared.Code);
        Assert.Equal(1, prepared.TrimmedLines);
    }

    [Fact]
    public void PrepareForWrite_StripsExportHeadersAndKeepsOptions()
    {
        var code = "Attribute VB_Name = \"Probe\"\r\nOption Explicit\r\n\r\nPublic Sub Probe()\r\nEnd Sub";

        var prepared = VbaProjectEditor.PrepareForWrite(code, "Option Compare Database\r\nOption Explicit");

        Assert.Equal("Option Explicit\r\n\r\nPublic Sub Probe()\r\nEnd Sub", prepared.Code);
    }

    [Fact]
    public void PrepareForWrite_KeepsTheExistingModulesOptionsWhenTheInputHasNone()
    {
        var prepared = VbaProjectEditor.PrepareForWrite(
            "Public Sub Probe()\r\nEnd Sub",
            "Option Compare Database\r\nOption Explicit\r\n\r\nPublic Sub Old()\r\nEnd Sub");

        Assert.Equal("Option Compare Database\r\nOption Explicit\r\n\r\nPublic Sub Probe()\r\nEnd Sub", prepared.Code);
    }

    [Fact]
    public void PrepareForWrite_DeduplicatesOptionLinesAndDropsTrailingBlankLines()
    {
        var prepared = VbaProjectEditor.PrepareForWrite(
            "Option Explicit\r\noption explicit\r\nPublic Sub Probe()\r\nEnd Sub\r\n\r\n\r\n",
            string.Empty);

        Assert.Equal("Option Explicit\r\n\r\nPublic Sub Probe()\r\nEnd Sub", prepared.Code);
    }

    [Fact]
    public void NormalizeForWrite_StripsExportHeadersAndKeepsOptions()
    {
        const string exported = """
Attribute VB_Name = "Module1"
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Option Compare Database
Option Explicit

Public Sub Hello()
End Sub
""";

        var result = VbaProjectEditor.NormalizeForWrite(exported, "Option Compare Database\r\nOption Explicit\r\n");

        Assert.DoesNotContain("Attribute", result);
        Assert.DoesNotContain("VERSION", result);
        Assert.DoesNotContain("BEGIN", result);
        Assert.StartsWith("Option Compare Database", result);
        Assert.Contains("Option Explicit", result);
        Assert.Contains("Public Sub Hello()", result);
    }

    [Fact]
    public void NormalizeForWrite_IsIdempotent()
    {
        var once = VbaProjectEditor.NormalizeForWrite("Attribute VB_Name = \"M\"\r\nOption Explicit\r\n\tSub A()\r\nEnd Sub  \r\n", string.Empty);
        var twice = VbaProjectEditor.NormalizeForWrite(once, string.Empty);

        Assert.Equal(once, twice);
    }

    [Theory]
    [InlineData(1023, false)]
    [InlineData(1024, true)]
    [InlineData(1502, true)]
    public void ValidateLineLengths_RefusesOnlyOverTheCodeModuleLimit(int length, bool refused)
    {
        var code = "Public Sub Probe()\r\n" + new string('a', length) + "\r\nEnd Sub";

        if (refused)
        {
            var exception = Assert.Throws<ModuleWriteRefusedException>(
                () => VbaProjectEditor.ValidateLineLengths(code, "Probe"));
            Assert.Contains("Line 2", exception.Message);
            Assert.Contains(length.ToString(), exception.Message);
            Assert.Contains("access_vba_import_file", exception.Message);
        }
        else
        {
            VbaProjectEditor.ValidateLineLengths(code, "Probe");
        }
    }

    [Fact]
    public void CompareCode_IgnoresLineEndingsAndTrailingWhitespace()
    {
        var intended = "Option Explicit\n\nPublic Sub Probe()\n    x = 1\nEnd Sub";
        var actual = "Option Explicit\r\n\r\nPublic Sub Probe()\r\n    x = 1   \r\nEnd Sub\r\n";

        var comparison = VbaProjectEditor.CompareCode(intended, actual);

        Assert.True(comparison.Same);
        Assert.False(comparison.CaseOnlyDifferences);
    }

    [Fact]
    public void CompareCode_AcceptsIdentifierCaseCorrection()
    {
        // The VBE re-cases identifiers to their declarations (e.g. IsArray -> isArray).
        var intended = "Public Sub Probe(IsArray As Boolean)\n    If IsArray Then End\nEnd Sub";
        var actual = "Public Sub Probe(isArray As Boolean)\n    If isArray Then End\nEnd Sub";

        var comparison = VbaProjectEditor.CompareCode(intended, actual);

        Assert.True(comparison.Same);
        Assert.True(comparison.CaseOnlyDifferences);
        Assert.Equal(2, comparison.CaseDifferenceLines);
    }

    [Fact]
    public void CompareCode_ReportsTheFirstRealDifference()
    {
        var intended = "Public Sub Probe()\n    x = 1\n    y = 2\nEnd Sub";
        var actual = "Public Sub Probe()\n    x = 1\n    y = 3\nEnd Sub";

        var comparison = VbaProjectEditor.CompareCode(intended, actual);

        Assert.False(comparison.Same);
        Assert.Equal(3, comparison.Line);
        Assert.Equal("    y = 2", comparison.Expected);
        Assert.Equal("    y = 3", comparison.Actual);

        var message = VbaProjectEditor.DescribeMismatch(comparison);
        Assert.Contains("line 3", message);
        Assert.Contains("y = 2", message);
        Assert.Contains("y = 3", message);
    }

    [Fact]
    public void CompareCode_ReportsExtraLinesFromTabExpansion()
    {
        // A tab passed to InsertLines becomes a line break, so the module gains a line.
        var intended = "Public Sub Probe()\n    x = 1\nEnd Sub";
        var actual = "Public Sub Probe()\n    \nx = 1\nEnd Sub";

        var comparison = VbaProjectEditor.CompareCode(intended, actual);

        Assert.False(comparison.Same);
        Assert.Equal(2, comparison.Line);
        Assert.Equal(3, comparison.ExpectedLines);
        Assert.Equal(4, comparison.ActualLines);
    }

    [Fact]
    public void FindProcedures_DetectsTestsAndSkips()
    {
        const string code = """
Option Explicit

'@test
Public Sub Passes()
End Sub

'@test-skip
Public Sub Skipped()
End Sub

'@test-xfail
Public Function ExpectedFailure() As Long
End Function

Private Sub NotATest()
End Sub

Public Property Get Value() As Long
End Property
""";

        var procedures = VbaProjectEditor.FindProcedures(code);

        Assert.Equal(5, procedures.Count);
        Assert.True(procedures.Single(p => p.Name == "Passes").IsTest);
        Assert.False(procedures.Single(p => p.Name == "Passes").IsSkipped);
        Assert.True(procedures.Single(p => p.Name == "Skipped").IsSkipped);
        Assert.True(procedures.Single(p => p.Name == "ExpectedFailure").IsExpectedFailure);
        Assert.False(procedures.Single(p => p.Name == "NotATest").IsPublic);
        Assert.Contains(procedures, p => p.Kind.StartsWith("Property", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void FindProcedures_ReportsOneBasedStartLinesAndNormalizedKinds()
    {
        const string code = "Option Explicit\r\nSub Implicit()\r\nEnd Sub\r\nFriend   Property   Let Value(v As Long)\r\nEnd Property\r\nPublic Static Function Counter() As Long\r\nEnd Function";

        var procedures = VbaProjectEditor.FindProcedures(code);

        var implicitSub = procedures.Single(p => p.Name == "Implicit");
        Assert.Equal(2, implicitSub.StartLine);
        Assert.True(implicitSub.IsPublic);
        Assert.Equal("Sub", implicitSub.Kind);

        var property = procedures.Single(p => p.Name == "Value");
        Assert.Equal("Property Let", property.Kind);
        Assert.False(property.IsPublic);

        Assert.Equal("Function", procedures.Single(p => p.Name == "Counter").Kind);
    }

    [Fact]
    public void FindProcedures_DirectiveSurvivesCommentsAndBlankLinesButNotCode()
    {
        const string code = """
'@test
' Explains the test.
Public Sub Commented()
End Sub

'@test

Public Sub AfterBlank()
End Sub

'@test
Private Const X = 1
Public Sub AfterCode()
End Sub
""";

        var procedures = VbaProjectEditor.FindProcedures(code);

        Assert.True(procedures.Single(p => p.Name == "Commented").IsTest);
        Assert.True(procedures.Single(p => p.Name == "AfterBlank").IsTest);
        Assert.False(procedures.Single(p => p.Name == "AfterCode").IsTest);
    }

    [Fact]
    public void FindProcedures_ADirectiveAfterABlankLineDoesNotOverrideThePendingOne()
    {
        // Once a blank line separates it from the directive, a later '@test-skip comment is just a comment.
        const string code = "'@test\r\n\r\n'@test-skip\r\nPublic Sub Probe()\r\nEnd Sub";

        var probe = Assert.Single(VbaProjectEditor.FindProcedures(code));

        Assert.True(probe.IsTest);
        Assert.False(probe.IsSkipped);
    }

    [Fact]
    public void FindProcedures_DirectiveIsCaseInsensitiveAndNeedsAWordBoundary()
    {
        const string code = "'@TEST\r\nSub Upper()\r\nEnd Sub\r\n'@testing\r\nSub NotADirective()\r\nEnd Sub";

        var procedures = VbaProjectEditor.FindProcedures(code);

        Assert.True(procedures.Single(p => p.Name == "Upper").IsTest);
        Assert.False(procedures.Single(p => p.Name == "NotADirective").IsTest);
    }

    [Fact]
    public void FindProcedures_IgnoresEndAndExitStatements()
    {
        var procedures = VbaProjectEditor.FindProcedures("Sub A()\r\n    Exit Sub\r\nEnd Sub\r\nFunction B()\r\n    Exit Function\r\nEnd Function");

        Assert.Equal(["A", "B"], procedures.Select(p => p.Name));
    }

    [Theory]
    [InlineData(VbaProjectEditor.StandardModule, "StandardModule")]
    [InlineData(VbaProjectEditor.ClassModule, "ClassModule")]
    [InlineData(VbaProjectEditor.FormModule, "FormModule")]
    [InlineData(VbaProjectEditor.DocumentModule, "DocumentModule")]
    [InlineData(11, "Type11")]
    public void GetTypeName_MapsComponentTypes(int type, string expected)
    {
        Assert.Equal(expected, VbaProjectEditor.GetTypeName(type));
    }

    [Fact]
    public void Sha256Hex_IgnoresLineEndingsAndTrailingWhitespace()
    {
        var crlf = VbaProjectEditor.Sha256Hex("Sub A()\r\n    x = 1  \r\nEnd Sub\r\n\r\n");
        var lf = VbaProjectEditor.Sha256Hex("Sub A()\n    x = 1\nEnd Sub");

        Assert.Equal(crlf, lf);
        Assert.Equal(64, lf.Length);
        Assert.Equal(lf.ToLowerInvariant(), lf);
    }

    [Fact]
    public void Sha256Hex_DistinguishesLeadingWhitespaceAndCase()
    {
        var baseline = VbaProjectEditor.Sha256Hex("Sub A()\n    x = 1\nEnd Sub");

        Assert.NotEqual(baseline, VbaProjectEditor.Sha256Hex("Sub A()\nx = 1\nEnd Sub"));
        Assert.NotEqual(baseline, VbaProjectEditor.Sha256Hex("Sub A()\n    X = 1\nEnd Sub"));
    }
}
