using DigitalJeeves.MsAccessMCP.Core;

namespace DigitalJeeves.MsAccessMCP.Tests;

public sealed class StartUpFormRecoveryTests : IDisposable
{
    private readonly TempDirectory _root = new();
    private readonly string _directory;

    public StartUpFormRecoveryTests()
    {
        // A subdirectory that does not exist yet, so Save has to create it.
        _directory = _root.Combine("recovery");
    }

    public void Dispose() => _root.Dispose();

    [Fact]
    public void SavedValueLoadsUntilCleared()
    {
        var recovery = new StartUpFormRecovery(_directory);
        var db = @"C:\Data\Sample.accdb";

        recovery.Save(db, "frmSwitchboard");
        Assert.Equal("frmSwitchboard", recovery.TryLoad(db));

        recovery.Clear(db);
        Assert.Null(recovery.TryLoad(db));
    }

    [Fact]
    public void PathMatchIgnoresCase()
    {
        var recovery = new StartUpFormRecovery(_directory);

        recovery.Save(@"C:\Data\Sample.accdb", "frmMain");

        Assert.Equal("frmMain", recovery.TryLoad(@"c:\data\SAMPLE.ACCDB"));
    }

    [Fact]
    public void MissingRecordOrDirectoryLoadsNull()
    {
        var recovery = new StartUpFormRecovery(_directory);

        Assert.Null(recovery.TryLoad(@"C:\Data\Sample.accdb"));
        recovery.Clear(@"C:\Data\Sample.accdb");
    }

    [Fact]
    public void RecordForAnotherDatabaseIsIgnored()
    {
        var recovery = new StartUpFormRecovery(_directory);
        var db = @"C:\Data\Sample.accdb";

        Directory.CreateDirectory(_directory);
        File.WriteAllLines(recovery.RecordPath(db), [@"C:\Data\Other.accdb", "frmOther"]);

        Assert.Null(recovery.TryLoad(db));
    }

    [Fact]
    public void TruncatedRecordLoadsNull()
    {
        var recovery = new StartUpFormRecovery(_directory);
        var db = @"C:\Data\Sample.accdb";

        Directory.CreateDirectory(_directory);
        File.WriteAllLines(recovery.RecordPath(db), [db]);

        Assert.Null(recovery.TryLoad(db));
    }

    [Fact]
    public void RecordsForDifferentDatabasesAreIndependent()
    {
        var recovery = new StartUpFormRecovery(_directory);
        var first = @"C:\Data\Sample.accdb";
        var second = @"C:\Data\Other.accdb";

        recovery.Save(first, "frmFirst");
        recovery.Save(second, "frmSecond");
        recovery.Clear(first);

        Assert.NotEqual(recovery.RecordPath(first), recovery.RecordPath(second));
        Assert.Null(recovery.TryLoad(first));
        Assert.Equal("frmSecond", recovery.TryLoad(second));
    }

    [Fact]
    public void RecordPathIsTheSameForEquivalentSpellings()
    {
        var recovery = new StartUpFormRecovery(_directory);

        Assert.Equal(
            recovery.RecordPath(@"C:\Data\Sample.accdb"),
            recovery.RecordPath(@"c:\data\sub\..\SAMPLE.accdb"));
    }
}
