using System.Globalization;
using System.IO;
using System.Text;

namespace DigitalJeeves.MsAccessMCP.Core;

/// <summary>
/// Per-call scratch workspace. The source database is copied here; every mutation (modules, kit,
/// compilation, MDE build) happens on the copy. The workspace is deleted unless the caller asks to
/// keep it, and a failed run keeps it regardless so the NDJSON and Access state can be inspected.
/// </summary>
internal sealed class VbaWorkspace : IDisposable
{
    private VbaWorkspace(string directory, string workDatabasePath, string? mdePath)
    {
        Directory = directory;
        WorkDatabasePath = workDatabasePath;
        MdePath = mdePath;
        ResultPath = Path.Combine(directory, "result.ndjson");
        LogPath = Path.Combine(directory, "run.log");
    }

    public string Directory { get; }
    public string WorkDatabasePath { get; }
    public string? MdePath { get; }
    public string ResultPath { get; }
    public string LogPath { get; }

    /// <summary>When true the directory is left on disk after the call.</summary>
    public bool Keep { get; set; }

    public bool Deleted { get; private set; }

    public IReadOnlyList<string> CleanupErrors { get; private set; } = [];

    public static VbaWorkspace Create(AccessMcpOptions options, string sourcePath, bool createMdePath)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(sourcePath);

        var root = AccessDbPathPolicy.NormalizeDirectory(
            options.VbaScratchDirectory,
            Path.Combine(Path.GetTempPath(), "AccessMCP.VbaTest"));

        EnsureWritable(root);

        var directory = Path.Combine(root, $"{DateTime.UtcNow.ToString("yyyyMMdd-HHmmss", CultureInfo.InvariantCulture)}-{Guid.NewGuid():N}");
        System.IO.Directory.CreateDirectory(directory);

        var extension = Path.GetExtension(sourcePath);
        var workPath = Path.Combine(directory, "work" + extension);
        CopyDatabase(sourcePath, workPath);

        string? mdePath = null;
        if (createMdePath)
        {
            mdePath = Path.Combine(directory, "out" + MdeBuilder.DefaultExtension(sourcePath));
        }

        return new VbaWorkspace(directory, workPath, mdePath);
    }

    public static VbaWorkspace CreateEmpty(AccessMcpOptions options, string extension, bool createMdePath)
    {
        var root = AccessDbPathPolicy.NormalizeDirectory(
            options.VbaScratchDirectory,
            Path.Combine(Path.GetTempPath(), "AccessMCP.VbaTest"));
        EnsureWritable(root);

        var directory = Path.Combine(root, $"{DateTime.UtcNow.ToString("yyyyMMdd-HHmmss", CultureInfo.InvariantCulture)}-{Guid.NewGuid():N}");
        System.IO.Directory.CreateDirectory(directory);

        var workPath = Path.Combine(directory, "work" + extension);
        var mdePath = createMdePath ? Path.Combine(directory, "out" + MdeBuilder.DefaultExtension(workPath)) : null;
        return new VbaWorkspace(directory, workPath, mdePath);
    }

    public void Log(string message)
    {
        try
        {
            File.AppendAllText(LogPath, $"{DateTime.UtcNow:O} {message}{Environment.NewLine}");
        }
        catch
        {
            // Logging must never fail a run.
        }
    }

    public void Dispose() => Cleanup(keepOnFailure: true, failed: false);

    public void Cleanup(bool keepOnFailure, bool failed)
    {
        if (Deleted)
        {
            return;
        }

        var keep = Keep || (keepOnFailure && failed);
        if (keep)
        {
            return;
        }

        var errors = new List<string>();
        try
        {
            foreach (var ldb in System.IO.Directory.EnumerateFiles(Directory, "*.ldb"))
            {
                try { File.Delete(ldb); } catch { }
            }

            System.IO.Directory.Delete(Directory, recursive: true);
            Deleted = true;
        }
        catch (Exception ex)
        {
            errors.Add(ex.Message);
        }

        CleanupErrors = errors;
    }

    private static void EnsureWritable(string directory)
    {
        System.IO.Directory.CreateDirectory(directory);
        var probe = Path.Combine(directory, $".write-{Guid.NewGuid():N}.tmp");
        File.WriteAllText(probe, "probe");
        File.Delete(probe);
    }

    /// <summary>
    /// Copies the source database. A lock file (or an exclusive open) means another process has the
    /// database open; copying then can produce a torn file, so the caller is told to close it.
    /// </summary>
    private static void CopyDatabase(string sourcePath, string destinationPath)
    {
        // A lock file is only a hint: it can be stale after a crash. The truth is whether the file
        // can be opened exclusively. Access keeps the database open while it finishes quitting, so
        // poll briefly instead of failing a run that follows immediately after a compile.
        if (!WaitForExclusiveRead(sourcePath, TimeSpan.FromSeconds(10)))
        {
            throw new InvalidOperationException(
                $"'{sourcePath}' is open (Access still holds it; its .ldb/.laccdb lock file exists). " +
                "Close it before running this tool.");
        }

        File.Copy(sourcePath, destinationPath, overwrite: true);

        var source = new FileInfo(sourcePath).Length;
        var destination = new FileInfo(destinationPath).Length;
        if (source != destination)
        {
            throw new IOException(
                $"The copy of '{sourcePath}' is inconsistent ({source} bytes in, {destination} bytes out). " +
                "The database is probably open; close it and retry.");
        }
    }

    private static bool WaitForExclusiveRead(string path, TimeSpan timeout)
    {
        var deadline = DateTime.UtcNow + timeout;
        while (true)
        {
            try
            {
                using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.None);
                return true;
            }
            catch (IOException) when (DateTime.UtcNow < deadline)
            {
                Thread.Sleep(250);
            }
            catch (IOException)
            {
                return false;
            }
            catch (UnauthorizedAccessException)
            {
                return false;
            }
        }
    }
}
