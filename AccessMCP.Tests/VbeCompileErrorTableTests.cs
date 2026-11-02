using DigitalJeeves.MsAccessMCP.Core;

namespace DigitalJeeves.MsAccessMCP.Tests;

public sealed class VbeCompileErrorTableTests
{
    [Fact]
    public void ContainsKnownCode()
    {
        Assert.True(VbeCompileErrorTable.TryGet(32807, out var message));
        Assert.Equal("User-defined type not defined", message);
    }

    [Theory]
    [InlineData(40079, "Expected End Sub")]
    [InlineData(40095, "Duplicate declaration in current scope")]
    [InlineData(40111, "Variable not defined")]
    [InlineData(32812, "Ambiguous name detected: <name>")]
    public void KnownCodesMapToMessages(int code, string expected)
    {
        Assert.True(VbeCompileErrorTable.TryGet(code, out var message));
        Assert.Equal(expected, message);
    }

    [Fact]
    public void UnknownCodeReturnsFalse()
    {
        Assert.False(VbeCompileErrorTable.TryGet(0x1234, out _));
    }

    [Fact]
    public void HasFullCompilerTable()
    {
        Assert.True(VbeCompileErrorTable.Count >= 150);
        Assert.False(string.IsNullOrWhiteSpace(VbeCompileErrorTable.SourceHash));
    }

    [Fact]
    public void VbeAddinClient_NullVbe_ReturnsNull()
    {
        Assert.Null(VbeAddinClient.TryConnect(null));
    }
}
