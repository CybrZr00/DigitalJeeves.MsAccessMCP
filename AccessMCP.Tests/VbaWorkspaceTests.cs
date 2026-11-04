using DigitalJeeves.MsAccessMCP.Core;

namespace DigitalJeeves.MsAccessMCP.Tests;

public sealed class VbaWorkspaceTests : IDisposable
{
    private readonly TempDirectory _dir = new();

    public void Dispose() => _dir.Dispose();

    [Fact]
    public void CreateCopiesTheSourceAndLeavesItUntouched()
    {
        var source = _dir.Touch("Sample.accdb", [1, 2, 3, 4, 5]);

        using var workspace = VbaWorkspace.Create(Options(), source, createMdePath: true);

        Assert.StartsWith(_dir.Combine("scratch"), workspace.Directory);
        Assert.Equal("work.accdb", Path.GetFileName(workspace.WorkDatabasePath));
        Assert.Equal([1, 2, 3, 4, 5], File.ReadAllBytes(workspace.WorkDatabasePath));
        Assert.Equal("out.accde", Path.GetFileName(workspace.MdePath));
        Assert.Equal(Path.Combine(workspace.Directory, "result.ndjson"), workspace.ResultPath);
        Assert.Equal([1, 2, 3, 4, 5], File.ReadAllBytes(source));
    }

    [Fact]
    public void MdbSourceGetsAnMdeOutputAndNoMdePathWhenNotRequested()
    {
        var source = _dir.Touch("Sample.mdb");

        using var withMde = VbaWorkspace.Create(Options(), source, createMdePath: true);
        using var withoutMde = VbaWorkspace.Create(Options(), source, createMdePath: false);

        Assert.Equal(".mde", Path.GetExtension(withMde.MdePath));
        Assert.Null(withoutMde.MdePath);
        Assert.NotEqual(withMde.Directory, withoutMde.Directory);
    }

    [Fact]
    public void CreateEmptyMakesTheDirectoryButNoDatabase()
    {
        using var workspace = VbaWorkspace.CreateEmpty(Options(), ".accdb", createMdePath: true);

        Assert.True(Directory.Exists(workspace.Directory));
        Assert.False(File.Exists(workspace.WorkDatabasePath));
        Assert.Equal(".accde", Path.GetExtension(workspace.MdePath));
    }

    [Fact]
    public void CleanupDeletesASuccessfulWorkspace()
    {
        var workspace = VbaWorkspace.Create(Options(), _dir.Touch("Sample.accdb"), createMdePath: false);
        workspace.Log("hello");
        File.WriteAllText(Path.Combine(workspace.Directory, "work.laccdb"), string.Empty);

        workspace.Cleanup(keepOnFailure: true, failed: false);

        Assert.True(workspace.Deleted);
        Assert.Empty(workspace.CleanupErrors);
        Assert.False(Directory.Exists(workspace.Directory));
    }

    [Fact]
    public void CleanupKeepsAFailedWorkspaceForInspection()
    {
        var workspace = VbaWorkspace.Create(Options(), _dir.Touch("Sample.accdb"), createMdePath: false);

        workspace.Cleanup(keepOnFailure: true, failed: true);

        Assert.False(workspace.Deleted);
        Assert.True(Directory.Exists(workspace.Directory));

        workspace.Cleanup(keepOnFailure: false, failed: true);
        Assert.True(workspace.Deleted);
    }

    [Fact]
    public void KeepOverridesCleanupAndDispose()
    {
        var workspace = VbaWorkspace.Create(Options(), _dir.Touch("Sample.accdb"), createMdePath: false);
        workspace.Keep = true;

        workspace.Cleanup(keepOnFailure: false, failed: false);
        workspace.Dispose();

        Assert.False(workspace.Deleted);
        Assert.True(Directory.Exists(workspace.Directory));
    }

    [Fact]
    public void LogAppendsTimestampedLines()
    {
        using var workspace = VbaWorkspace.CreateEmpty(Options(), ".accdb", createMdePath: false);

        workspace.Log("first");
        workspace.Log("second");

        var lines = File.ReadAllLines(workspace.LogPath);
        Assert.Equal(2, lines.Length);
        Assert.EndsWith(" first", lines[0]);
        Assert.EndsWith(" second", lines[1]);
    }

    [Fact]
    public void CreateRefusesASourceThatIsOpenExclusively()
    {
        var source = _dir.Touch("Locked.accdb");
        using var lockStream = new FileStream(source, FileMode.Open, FileAccess.ReadWrite, FileShare.None);

        var exception = Assert.Throws<InvalidOperationException>(
            () => VbaWorkspace.Create(Options(), source, createMdePath: false));

        Assert.Contains("is open", exception.Message);
    }

    [Fact]
    public void CreateInPlaceWorksOnTheSourceAndMakesNoCopy()
    {
        var source = _dir.Touch("Sample.accdb", [1, 2, 3]);

        using var workspace = VbaWorkspace.CreateInPlace(Options(), source, createMdePath: true);

        Assert.True(workspace.InPlace);
        Assert.Equal(Path.GetFullPath(source), workspace.WorkDatabasePath);
        Assert.True(Directory.Exists(workspace.Directory));
        Assert.StartsWith(_dir.Combine("scratch"), workspace.Directory);
        Assert.Equal("out.accde", Path.GetFileName(workspace.MdePath));
        Assert.Equal([1, 2, 3], File.ReadAllBytes(source));
    }

    [Fact]
    public void CreateInPlaceCleanupDeletesOnlyTheArtifactDirectory()
    {
        var source = _dir.Touch("Sample.accdb", [9, 9, 9]);
        var workspace = VbaWorkspace.CreateInPlace(Options(), source, createMdePath: false);
        workspace.Log("hello");

        workspace.Cleanup(keepOnFailure: false, failed: false);

        Assert.True(workspace.Deleted);
        Assert.False(Directory.Exists(workspace.Directory));
        Assert.True(File.Exists(source));
        Assert.Equal([9, 9, 9], File.ReadAllBytes(source));
    }

    [Fact]
    public void CreateInPlaceDoesNotRequireExclusiveAccessToTheSource()
    {
        // The point of in-place is to let Access open that very file, so the exclusive-open
        // probe used by Create must not apply.
        var source = _dir.Touch("Locked.accdb");
        using var lockStream = new FileStream(source, FileMode.Open, FileAccess.ReadWrite, FileShare.None);

        using var workspace = VbaWorkspace.CreateInPlace(Options(), source, createMdePath: false);

        Assert.Equal(Path.GetFullPath(source), workspace.WorkDatabasePath);
    }

    [Fact]
    public void CreateInPlaceRejectsAMissingDatabase()
    {
        var missing = _dir.Combine("Nope.accdb");

        var exception = Assert.Throws<FileNotFoundException>(
            () => VbaWorkspace.CreateInPlace(Options(), missing, createMdePath: false));

        Assert.Contains("Nope.accdb", exception.Message);
    }

    private AccessMcpOptions Options() => new() { VbaScratchDirectory = _dir.Combine("scratch") };
}
