using DigitalJeeves.MsAccessMCP.Core;

namespace DigitalJeeves.MsAccessMCP.Tests;

public sealed class VbaMdeCacheTests : IDisposable
{
    private readonly TempDirectory _dir = new();

    public void Dispose() => _dir.Dispose();

    [Fact]
    public void CacheKey_IsStableAndChangesWithModuleHash()
    {
        var source = _dir.Touch("source.accdb", [1, 2, 3, 4]);
        var modules = new[] { new VbaModuleSpec("M", "standard", "Public Sub A()\nEnd Sub", null) };

        var first = Key(source, modules);
        var second = Key(source, modules);
        var changed = Key(source, [new VbaModuleSpec("M", "standard", "Public Sub B()\nEnd Sub", null)]);

        Assert.Equal(first, second);
        Assert.NotEqual(first, changed);
        Assert.Equal(64, first.Length);
    }

    [Fact]
    public void CacheKey_IgnoresModuleAndTestOrder()
    {
        var source = _dir.Touch("source.accdb", [1, 2, 3, 4]);
        var a = new VbaModuleSpec("A", "standard", "Sub A()\nEnd Sub", null);
        var b = new VbaModuleSpec("B", "class", "Sub B()\nEnd Sub", null);

        Assert.Equal(
            Key(source, [a, b], tests: ["T1", "T2"]),
            Key(source, [b, a], tests: ["T2", "T1"]));
    }

    [Fact]
    public void CacheKey_ChangesWithEveryOtherInput()
    {
        var source = _dir.Touch("source.accdb", [1, 2, 3, 4]);
        var modules = new[] { new VbaModuleSpec("M", "standard", "Sub A()\nEnd Sub", null) };
        var baseline = Key(source, modules);

        Assert.NotEqual(baseline, Key(source, modules, kit: "other-kit"));
        Assert.NotEqual(baseline, Key(source, modules, decompile: true));
        Assert.NotEqual(baseline, Key(source, modules, version: "16.1"));
        Assert.NotEqual(baseline, Key(source, modules, bitness: "x64"));
        Assert.NotEqual(baseline, Key(source, modules, entryPoint: "Other"));
        Assert.NotEqual(baseline, Key(source, modules, tests: ["A", "B"]));
        Assert.NotEqual(baseline, Key(source, [new VbaModuleSpec("M", "class", "Sub A()\nEnd Sub", null)]));
    }

    [Fact]
    public void CacheKey_ChangesWhenTheSourceDatabaseChanges()
    {
        var source = _dir.Touch("source.accdb", [1, 2, 3, 4]);
        var modules = new[] { new VbaModuleSpec("M", "standard", "Sub A()\nEnd Sub", null) };
        var before = Key(source, modules);

        File.WriteAllBytes(source, [1, 2, 3, 4, 5]);

        Assert.NotEqual(before, Key(source, modules));
    }

    [Fact]
    public void PutThenTryGetReturnsTheCachedCopy()
    {
        var cache = new VbaMdeCache(Options());
        var built = _dir.Touch("built.accde", [9, 9, 9]);
        var key = new string('a', 64);

        cache.Put(key, built, Manifest(key));

        Assert.True(cache.TryGet(key, out var path, out var problem));
        Assert.Null(problem);
        Assert.Equal([9, 9, 9], File.ReadAllBytes(path));
        Assert.StartsWith(cache.Root, path);
        Assert.Equal(".accde", Path.GetExtension(path));
    }

    [Fact]
    public void TryGetMissesForUnknownKeyAndReportsAnEmptyFile()
    {
        var cache = new VbaMdeCache(Options());
        var key = new string('b', 64);

        Assert.False(cache.TryGet(key, out _, out var none));
        Assert.Null(none);

        cache.Put(key, _dir.Touch("empty.accde", []), Manifest(key));

        Assert.False(cache.TryGet(key, out _, out var problem));
        Assert.Contains("missing or empty", problem);
    }

    [Fact]
    public void TryGetTreatsACorruptManifestAsAMiss()
    {
        var cache = new VbaMdeCache(Options());
        var key = new string('c', 64);
        cache.Put(key, _dir.Touch("built.accde", [1]), Manifest(key));

        File.WriteAllText(Path.Combine(cache.Root, "entries", key, "manifest.json"), "{ not json");

        Assert.False(cache.TryGet(key, out _, out var problem));
        Assert.NotNull(problem);
    }

    [Fact]
    public void DisabledCacheNeitherStoresNorReturns()
    {
        var cache = new VbaMdeCache(Options() with { VbaCacheEnabled = false });
        var key = new string('d', 64);

        cache.Put(key, _dir.Touch("built.accde", [1]), Manifest(key));

        Assert.False(cache.TryGet(key, out _, out _));
        Assert.Empty(cache.Info().Entries);
        Assert.False(cache.Info().Enabled);
    }

    [Fact]
    public void InfoListsEntriesNewestFirstWithTotals()
    {
        var cache = new VbaMdeCache(Options());
        var older = new string('1', 64);
        var newer = new string('2', 64);

        cache.Put(older, _dir.Touch("a.accde", [1, 1]), Manifest(older, DateTimeOffset.UtcNow.AddHours(-1)));
        cache.Put(newer, _dir.Touch("b.accde", [2, 2, 2]), Manifest(newer, DateTimeOffset.UtcNow));

        var info = cache.Info();

        Assert.Equal([newer, older], info.Entries.Select(e => e.Key));
        Assert.Equal(5, info.TotalBytes);
        Assert.Equal(cache.Root, info.Root);
    }

    [Fact]
    public void ClearRemovesOneKeyOrEverything()
    {
        var cache = new VbaMdeCache(Options());
        var first = new string('e', 64);
        var second = new string('f', 64);
        cache.Put(first, _dir.Touch("a.accde", [1, 2]), Manifest(first));
        cache.Put(second, _dir.Touch("b.accde", [1, 2]), Manifest(second));

        var (removedOne, freedOne) = cache.Clear(first, all: false);
        Assert.Equal(1, removedOne);
        Assert.True(freedOne >= 2);
        Assert.Equal([second], cache.Info().Entries.Select(e => e.Key));

        Assert.Equal((0, 0L), cache.Clear(first, all: false));
        Assert.Equal((0, 0L), cache.Clear(null, all: false));

        var (removedAll, _) = cache.Clear(null, all: true);
        Assert.Equal(1, removedAll);
        Assert.Empty(cache.Info().Entries);
    }

    [Fact]
    public void ClearNeverDeletesOutsideTheCacheRoot()
    {
        // The key comes straight from the access_vba_cache_clear caller. This one resolves to a sibling
        // of the cache root, still inside the test's temp directory.
        const string key = @"..\..\victim";
        var cache = new VbaMdeCache(Options());
        Directory.CreateDirectory(Path.Combine(cache.Root, "entries"));
        var victim = Path.GetFullPath(Path.Combine(cache.Root, "entries", key));
        Directory.CreateDirectory(victim);
        File.WriteAllText(Path.Combine(victim, "keep.txt"), "not part of the cache");

        Assert.Throws<ArgumentException>(() => cache.Clear(key, all: false));
        Assert.True(File.Exists(Path.Combine(victim, "keep.txt")));
    }

    [Theory]
    [InlineData("abc")]
    [InlineData("entries")]
    [InlineData("gggggggggggggggggggggggggggggggggggggggggggggggggggggggggggggggg")]
    public void ClearRejectsAnythingThatIsNotAKey(string key)
    {
        var cache = new VbaMdeCache(Options());

        Assert.Throws<ArgumentException>(() => cache.Clear(key, all: false));
    }

    [Fact]
    public void ClearAcceptsUpperCaseKeys()
    {
        var cache = new VbaMdeCache(Options());
        var key = new string('a', 64);
        cache.Put(key, _dir.Touch("built.accde", [1]), Manifest(key));

        Assert.Equal(1, cache.Clear(key.ToUpperInvariant(), all: false).Removed);
    }

    [Fact]
    public void PutTrimsOldestEntriesOverTheSizeCap()
    {
        var cache = new VbaMdeCache(Options() with { VbaCacheMaxBytes = 5 });
        var oldest = new string('7', 64);
        var newest = new string('8', 64);

        cache.Put(oldest, _dir.Touch("a.accde", [1, 1, 1, 1]), Manifest(oldest, DateTimeOffset.UtcNow.AddHours(-1)));
        cache.Put(newest, _dir.Touch("b.accde", [2, 2, 2, 2]), Manifest(newest, DateTimeOffset.UtcNow));

        Assert.Equal([newest], cache.Info().Entries.Select(e => e.Key));
    }

    [Fact]
    public void RootDefaultsUnderLocalAppData()
    {
        var cache = new VbaMdeCache(new AccessMcpOptions());

        Assert.StartsWith(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            cache.Root,
            StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ModuleSpec_KindIsNormalized()
    {
        Assert.Equal("class", new VbaModuleSpec("M", " CLS ", "", null).StableKind);
        Assert.Equal("class", new VbaModuleSpec("M", "Class", "", null).StableKind);
        Assert.Equal("standard", new VbaModuleSpec("M", "anything", "", null).StableKind);
    }

    [Fact]
    public void ModuleSpec_ResolvesCodeInlineFromFileOrEmpty()
    {
        var file = _dir.Combine("M.bas");
        File.WriteAllText(file, "Sub FromFile()\r\nEnd Sub");

        Assert.Equal("inline", new VbaModuleSpec("M", "standard", "inline", file).ResolveCode());
        Assert.Equal("Sub FromFile()\r\nEnd Sub", new VbaModuleSpec("M", "standard", null, file).ResolveCode());
        Assert.Equal(string.Empty, new VbaModuleSpec("M", "standard", null, null).ResolveCode());
        Assert.Throws<FileNotFoundException>(() => new VbaModuleSpec("M", "standard", null, _dir.Combine("missing.bas")).ResolveCode());
    }

    [Fact]
    public void ModuleSpec_HashIgnoresExportHeaders()
    {
        var plain = new VbaModuleSpec("M", "standard", "Sub A()\r\nEnd Sub", null);
        var exported = new VbaModuleSpec("M", "standard", "Attribute VB_Name = \"M\"\r\nSub A()\r\nEnd Sub\r\n", null);

        Assert.Equal(plain.Sha256, exported.Sha256);
    }

    private AccessMcpOptions Options() => new() { VbaCacheDirectory = _dir.Combine("cache") };

    private static string Key(
        string source,
        IReadOnlyList<VbaModuleSpec> modules,
        string kit = "kit",
        bool decompile = false,
        string version = "16.0",
        string bitness = "x86",
        string entryPoint = "RunAllTests",
        IReadOnlyList<string>? tests = null)
        => VbaMdeCache.ComputeKey(source, modules, kit, decompile, version, bitness, entryPoint, tests ?? ["A"]);

    private static VbaCacheManifest Manifest(string key, DateTimeOffset? createdAt = null) => new(
        key,
        "accdb",
        @"C:\Data\Sample.accdb",
        0,
        0,
        "16.0",
        "x64",
        "kit",
        false,
        "RunAllTests",
        [],
        [],
        createdAt ?? DateTimeOffset.UtcNow);
}
