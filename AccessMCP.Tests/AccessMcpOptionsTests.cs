using DigitalJeeves.MsAccessMCP.Core;

namespace DigitalJeeves.MsAccessMCP.Tests;

public sealed class AccessMcpOptionsTests
{
    [Fact]
    public void DefaultsUseExpectedProvider()
    {
        var options = new AccessMcpOptions();

        Assert.Equal("Microsoft.ACE.OLEDB.16.0", options.OleDbProvider);
    }

    [Fact]
    public void DefaultsAreReadWriteWithNoDefaultDatabase()
    {
        var options = new AccessMcpOptions();

        Assert.False(options.ReadOnly);
        Assert.False(options.EnableComAutomation);
        Assert.Null(options.DefaultDatabasePath);
    }

    [Fact]
    public void DefaultsUseExpectedCommandTimeout()
    {
        var options = new AccessMcpOptions();

        Assert.Equal(30, options.CommandTimeoutSeconds);
    }

    [Fact]
    public void DefaultsSuppressStartupForm()
    {
        var options = new AccessMcpOptions();

        Assert.True(options.SuppressStartupForm);
    }

    [Fact]
    public void VbaDefaultsAreSafe()
    {
        var options = new AccessMcpOptions();

        Assert.False(options.EnableVbaTestTools);
        Assert.True(options.VbaCacheEnabled);
        Assert.Equal(120, options.VbaTimeoutSeconds);
        Assert.Equal("syscmd", options.VbaMdeBuildStrategy);
    }

    [Fact]
    public void RemainingDefaultsMatchTheDocumentedValues()
    {
        var options = new AccessMcpOptions();

        Assert.Equal("Access.Application", options.AccessComProgId);
        Assert.Equal("DAO.DBEngine.120", options.DaoProgId);
        Assert.Null(options.WorkgroupPath);
        Assert.Null(options.VbaScratchDirectory);
        Assert.Null(options.VbaCacheDirectory);
        Assert.Equal(60, options.VbaCompileTimeoutSeconds);
        Assert.False(options.VbaDecompileFirst);
        Assert.Equal(2L * 1024 * 1024 * 1024, options.VbaCacheMaxBytes);
        Assert.False(options.KeepVbaArtifacts);
        Assert.Equal(200_000, options.VbaMaxResultChars);
        Assert.Equal(10, options.VbaCleanupTimeoutSeconds);
    }

    [Fact]
    public void OptionsAreAValueRecord()
    {
        // Hosts compare and copy options with `with`; a reference-type rewrite would break that.
        var original = new AccessMcpOptions { DefaultDatabasePath = @"C:\Data\Sample.accdb" };
        var copy = original with { };
        var readOnly = original with { ReadOnly = true };

        Assert.Equal(original, copy);
        Assert.NotEqual(original, readOnly);
        Assert.False(original.ReadOnly);
    }
}
