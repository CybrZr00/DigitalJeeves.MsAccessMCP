using DigitalJeeves.MsAccessMCP.Core;

namespace DigitalJeeves.MsAccessMCP.Tests;

public sealed class AccessDbPathPolicyTests
{
    [Fact]
    public void EditableAndCompiledExtensions()
    {
        using var dir = new TempDirectory();
        var mdb = dir.Touch("a.mdb");
        var accdb = dir.Touch("a.accdb");
        var mde = dir.Touch("a.mde");
        var accde = dir.Touch("a.accde");
        var options = new AccessMcpOptions { DefaultDatabasePath = mdb };

        Assert.Equal(mdb, AccessDbPathPolicy.NormalizeAndValidate(mdb, options));
        Assert.Equal(accdb, AccessDbPathPolicy.NormalizeAndValidate(accdb, options));
        Assert.Throws<ArgumentException>(() => AccessDbPathPolicy.NormalizeAndValidate(mde, options));

        Assert.Equal(mde, AccessDbPathPolicy.NormalizeComDatabasePath(mde, options));
        Assert.Equal(accde, AccessDbPathPolicy.NormalizeComDatabasePath(accde, options));
        Assert.Equal("mde", AccessDbPathPolicy.GetDatabaseKind(mde));
        Assert.True(AccessDbPathPolicy.IsCompiledDatabase(accde));
        Assert.False(AccessDbPathPolicy.IsCompiledDatabase(accdb));
    }

    [Fact]
    public void ComPathRefusesCompiledDatabasesWhenNotAllowed()
    {
        using var dir = new TempDirectory();
        var mde = dir.Touch("a.mde");

        var exception = Assert.Throws<ArgumentException>(
            () => AccessDbPathPolicy.NormalizeComDatabasePath(mde, new AccessMcpOptions(), allowCompiled: false));

        Assert.DoesNotContain(".accde", exception.Message);
    }

    [Theory]
    [InlineData("a.txt")]
    [InlineData("a.laccdb")]
    [InlineData("a")]
    public void OtherExtensionsAreRefusedByBothChecks(string name)
    {
        using var dir = new TempDirectory();
        var path = dir.Touch(name);
        var options = new AccessMcpOptions();

        Assert.Throws<ArgumentException>(() => AccessDbPathPolicy.NormalizeAndValidate(path, options));
        Assert.Throws<ArgumentException>(() => AccessDbPathPolicy.NormalizeComDatabasePath(path, options));
    }

    [Fact]
    public void ExtensionCheckIgnoresCase()
    {
        using var dir = new TempDirectory();
        var path = dir.Touch("Upper.ACCDB");

        Assert.Equal(path, AccessDbPathPolicy.NormalizeAndValidate(path, new AccessMcpOptions()));
        Assert.Equal("accdb", AccessDbPathPolicy.GetDatabaseKind(path));
        Assert.True(AccessDbPathPolicy.IsCompiledDatabase("X.MDE"));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void FallsBackToDefaultDatabasePath(string? databasePath)
    {
        using var dir = new TempDirectory();
        var db = dir.Touch("default.accdb");
        var options = new AccessMcpOptions { DefaultDatabasePath = db };

        Assert.Equal(db, AccessDbPathPolicy.NormalizeAndValidate(databasePath, options));
    }

    [Fact]
    public void ExplicitPathWinsOverDefault()
    {
        using var dir = new TempDirectory();
        var fallback = dir.Touch("default.accdb");
        var explicitPath = dir.Touch("explicit.accdb");
        var options = new AccessMcpOptions { DefaultDatabasePath = fallback };

        Assert.Equal(explicitPath, AccessDbPathPolicy.NormalizeAndValidate(explicitPath, options));
    }

    [Fact]
    public void MissingPathWithNoDefaultIsAnArgumentError()
    {
        var exception = Assert.Throws<ArgumentException>(
            () => AccessDbPathPolicy.NormalizeAndValidate(null, new AccessMcpOptions()));

        Assert.Equal("databasePath", exception.ParamName);
    }

    [Fact]
    public void NonexistentFileIsReportedWithItsFullPath()
    {
        using var dir = new TempDirectory();
        var missing = dir.Combine("missing.accdb");

        var exception = Assert.Throws<FileNotFoundException>(
            () => AccessDbPathPolicy.NormalizeAndValidate(missing, new AccessMcpOptions()));

        Assert.Equal(missing, exception.FileName);
    }

    [Fact]
    public void RelativePathsAndEnvironmentVariablesAreExpanded()
    {
        using var dir = new TempDirectory();
        var db = dir.Touch("env.accdb");
        var variable = $"ACCESSMCP_TEST_{Guid.NewGuid():N}";
        Environment.SetEnvironmentVariable(variable, dir.Path);
        try
        {
            Assert.Equal(db, AccessDbPathPolicy.NormalizeAndValidate($"%{variable}%\\env.accdb", new AccessMcpOptions()));
            Assert.Equal(db, AccessDbPathPolicy.NormalizeAndValidate(Path.Combine(dir.Path, "sub", "..", "env.accdb"), new AccessMcpOptions()));
        }
        finally
        {
            Environment.SetEnvironmentVariable(variable, null);
        }
    }

    [Fact]
    public void NormalizeDirectoryUsesFallbackOnlyWhenNothingIsConfigured()
    {
        using var dir = new TempDirectory();
        var fallback = dir.Combine("fallback");
        var configured = dir.Combine("configured");

        Assert.Equal(fallback, AccessDbPathPolicy.NormalizeDirectory(null, fallback));
        Assert.Equal(fallback, AccessDbPathPolicy.NormalizeDirectory("  ", fallback));
        Assert.Equal(configured, AccessDbPathPolicy.NormalizeDirectory(configured, fallback));
    }
}
