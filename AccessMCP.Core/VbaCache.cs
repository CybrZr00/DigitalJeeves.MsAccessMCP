using System.Globalization;
using System.IO;
using System.Text;
using System.Text.Json;

namespace DigitalJeeves.MsAccessMCP.Core;

/// <summary>A module as it participated in a cache key.</summary>
internal sealed record VbaCachedModule(string Name, string Kind, string Sha256, int Lines);

internal sealed record VbaCacheManifest(
    string Key,
    string Kind,
    string SourcePath,
    long SourceSize,
    long SourceWriteTicks,
    string AccessVersion,
    string AccessBitness,
    string KitSha256,
    bool Decompile,
    string EntryPoint,
    IReadOnlyList<string> Tests,
    IReadOnlyList<VbaCachedModule> Modules,
    DateTimeOffset CreatedAt);

internal sealed record VbaCacheEntry(string Key, string Path, long Bytes, DateTimeOffset CreatedAt);

internal sealed record VbaCacheInfo(string Root, bool Enabled, IReadOnlyList<VbaCacheEntry> Entries, long TotalBytes, long MaxBytes);

/// <summary>
/// Cache for built MDE/ACCDE files. Building is the slow part of an MDE scenario; the prepared MDB
/// is always materialised per run, so only the compiled output is cached. Entries are keyed on the
/// Access build/bitness, source identity, every injected module hash, the kit hash, the decompile
/// flag and the generated test manifest.
/// </summary>
internal sealed class VbaMdeCache
{
    private const int SchemaVersion = 1;
    private static readonly TimeSpan LockTimeout = TimeSpan.FromMinutes(5);
    private static readonly JsonSerializerOptions JsonOptions = new(JsonSerializerDefaults.Web) { WriteIndented = true };

    private readonly AccessMcpOptions _options;

    public VbaMdeCache(AccessMcpOptions options)
    {
        _options = options;
        Root = AccessDbPathPolicy.NormalizeDirectory(
            options.VbaCacheDirectory,
            Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "AccessMCP",
                "VbaTest",
                "cache"));
        Enabled = options.VbaCacheEnabled;
        MaxBytes = options.VbaCacheMaxBytes;
    }

    public string Root { get; }
    public bool Enabled { get; }
    public long MaxBytes { get; }

    /// <summary>
    /// Computes the cache key. Every input that can change the compiled output is included; the
    /// source's content is represented by its size and last-write ticks rather than a full hash of
    /// a potentially multi-hundred-megabyte database.
    /// </summary>
    public static string ComputeKey(
        string sourcePath,
        IReadOnlyList<VbaModuleSpec> modules,
        string kitSha256,
        bool decompile,
        string accessVersion,
        string accessBitness,
        string entryPoint,
        IReadOnlyList<string> tests)
    {
        var source = new FileInfo(sourcePath);
        var builder = new StringBuilder();
        builder.Append("schema=").Append(SchemaVersion).Append('\n');
        builder.Append("access=").Append(accessVersion).Append('/').Append(accessBitness).Append('\n');
        builder.Append("source=").Append(Path.GetFullPath(sourcePath)).Append('\n');
        builder.Append("size=").Append(source.Length.ToString(CultureInfo.InvariantCulture)).Append('\n');
        builder.Append("mtime=").Append(source.LastWriteTimeUtc.Ticks.ToString(CultureInfo.InvariantCulture)).Append('\n');
        builder.Append("kit=").Append(kitSha256).Append('\n');
        builder.Append("decompile=").Append(decompile).Append('\n');
        builder.Append("entry=").Append(entryPoint).Append('\n');

        foreach (var module in modules.OrderBy(m => m.Name, StringComparer.OrdinalIgnoreCase))
        {
            builder.Append("module:").Append(module.Name).Append(':').Append(module.Kind).Append('=').Append(module.Sha256).Append('\n');
        }

        foreach (var test in tests.OrderBy(t => t, StringComparer.OrdinalIgnoreCase))
        {
            builder.Append("test:").Append(test).Append('\n');
        }

        return VbaProjectEditor.Sha256Hex(builder.ToString());
    }

    /// <summary>
    /// Returns the cached compiled database, or false. A corrupt manifest is treated as a miss.
    /// </summary>
    public bool TryGet(string key, out string mdePath, out string? problem)
    {
        mdePath = string.Empty;
        problem = null;

        if (!Enabled)
        {
            return false;
        }

        var directory = EntryDirectory(key);
        var manifestPath = Path.Combine(directory, "manifest.json");
        if (!File.Exists(manifestPath))
        {
            return false;
        }

        try
        {
            using var stream = File.OpenRead(manifestPath);
            var manifest = JsonSerializer.Deserialize<VbaCacheManifest>(stream, JsonOptions);
            if (manifest is null || !manifest.Key.Equals(key, StringComparison.OrdinalIgnoreCase))
            {
                problem = "manifest key mismatch";
                return false;
            }

            var extension = manifest.Kind.Contains("accdb", StringComparison.OrdinalIgnoreCase) ? ".accde" : ".mde";
            var candidate = Path.Combine(directory, "base" + extension);
            if (!File.Exists(candidate) || new FileInfo(candidate).Length == 0)
            {
                problem = "cached file is missing or empty";
                return false;
            }

            File.SetLastAccessTimeUtc(candidate, DateTime.UtcNow);
            mdePath = candidate;
            return true;
        }
        catch (Exception ex)
        {
            problem = ex.Message;
            return false;
        }
    }

    /// <summary>Stores a freshly built file under <paramref name="key"/>.</summary>
    public void Put(string key, string builtPath, VbaCacheManifest manifest)
    {
        if (!Enabled)
        {
            return;
        }

        var directory = EntryDirectory(key);
        Directory.CreateDirectory(directory);

        using var lockStream = AcquireLock(key);
        var extension = Path.GetExtension(builtPath);
        var target = Path.Combine(directory, "base" + extension);
        var temp = target + ".tmp-" + Guid.NewGuid().ToString("N");

        try
        {
            File.Copy(builtPath, temp, overwrite: true);
            File.Move(temp, target, overwrite: true);
            File.WriteAllText(Path.Combine(directory, "manifest.json"), JsonSerializer.Serialize(manifest, JsonOptions));
            File.SetLastAccessTimeUtc(target, DateTime.UtcNow);
        }
        finally
        {
            try { if (File.Exists(temp)) File.Delete(temp); } catch { }
        }

        Trim();
    }

    public VbaCacheInfo Info()
    {
        var entries = new List<VbaCacheEntry>();
        long total = 0;

        foreach (var directory in EnumerateEntryDirectories())
        {
            var manifestPath = Path.Combine(directory, "manifest.json");
            VbaCacheManifest? manifest = null;
            try
            {
                using var stream = File.OpenRead(manifestPath);
                manifest = JsonSerializer.Deserialize<VbaCacheManifest>(stream, JsonOptions);
            }
            catch
            {
            }

            var file = Directory.EnumerateFiles(directory, "base.*").FirstOrDefault();
            var size = file is null ? 0 : new FileInfo(file).Length;
            total += size;

            entries.Add(new VbaCacheEntry(
                Path.GetFileName(directory),
                file ?? directory,
                size,
                manifest?.CreatedAt ?? DateTimeOffset.MinValue));
        }

        return new VbaCacheInfo(Root, Enabled, entries.OrderByDescending(e => e.CreatedAt).ToArray(), total, MaxBytes);
    }

    public (int Removed, long Freed) Clear(string? key, bool all)
    {
        // The key arrives from the MCP caller and becomes a directory name. Refuse anything that isn't a
        // key ComputeKey could have produced, so "..\" can never reach outside the cache.
        if (!all && key is not null && !IsCacheKey(key))
        {
            throw new ArgumentException("Not a cache key. Keys are 64 hexadecimal characters, as listed by access_vba_cache_info.", nameof(key));
        }

        var removed = 0;
        long freed = 0;

        var directories = all
            ? EnumerateEntryDirectories().ToArray()
            : key is null
                ? []
                : Directory.Exists(EntryDirectory(key)) ? [EntryDirectory(key)] : [];

        foreach (var directory in directories)
        {
            try
            {
                foreach (var file in Directory.EnumerateFiles(directory, "*", SearchOption.AllDirectories))
                {
                    try { freed += new FileInfo(file).Length; } catch { }
                }

                Directory.Delete(directory, recursive: true);
                removed++;
            }
            catch
            {
                // Leave entries we cannot delete; the report still describes what worked.
            }
        }

        return (removed, freed);
    }

    private void Trim()
    {
        if (MaxBytes <= 0)
        {
            return;
        }

        var info = Info();
        var total = info.TotalBytes;
        foreach (var entry in info.Entries.OrderBy(e => e.CreatedAt))
        {
            if (total <= MaxBytes)
            {
                break;
            }

            try
            {
                Directory.Delete(Path.GetDirectoryName(entry.Path)!, recursive: true);
                total -= entry.Bytes;
            }
            catch
            {
            }
        }
    }

    private IEnumerable<string> EnumerateEntryDirectories()
    {
        var entries = Path.Combine(Root, "entries");
        if (!Directory.Exists(entries))
        {
            yield break;
        }

        foreach (var directory in Directory.EnumerateDirectories(entries))
        {
            yield return directory;
        }
    }

    private string EntryDirectory(string key) => Path.Combine(Root, "entries", key);

    private static bool IsCacheKey(string key) => key.Length == 64 && key.All(char.IsAsciiHexDigit);

    private FileStream AcquireLock(string key)
    {
        var lockDirectory = Path.Combine(Root, "locks");
        Directory.CreateDirectory(lockDirectory);
        var lockPath = Path.Combine(lockDirectory, key + ".lock");

        var deadline = DateTime.UtcNow + LockTimeout;
        while (true)
        {
            try
            {
                return new FileStream(lockPath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
            }
            catch (IOException) when (DateTime.UtcNow < deadline)
            {
                Thread.Sleep(200);
            }
        }
    }
}
