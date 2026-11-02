using System.Runtime.InteropServices;

namespace DigitalJeeves.MsAccessMCP.Tests;

/// <summary>A scratch directory under %TEMP% that is deleted when the test finishes.</summary>
internal sealed class TempDirectory : IDisposable
{
    public TempDirectory()
    {
        Path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "AccessMCP.Tests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(Path);
    }

    public string Path { get; }

    public string Combine(string name) => System.IO.Path.Combine(Path, name);

    /// <summary>Writes a placeholder file. Enough for path policy checks; not a real database.</summary>
    public string Touch(string name, byte[]? content = null)
    {
        var path = Combine(name);
        File.WriteAllBytes(path, content ?? [0]);
        return path;
    }

    public void Dispose()
    {
        try { Directory.Delete(Path, recursive: true); } catch (IOException) { } catch (UnauthorizedAccessException) { }
    }
}

/// <summary>
/// A fact that needs the ACE OLE DB provider (and ADOX, to create the database). Skipped on machines
/// without it, or when the test process bitness doesn't match the installed ACE.
/// </summary>
internal sealed class AceFactAttribute : FactAttribute
{
    public AceFactAttribute()
    {
        if (!TestDatabase.AceAvailable)
        {
            Skip = "ACE OLE DB provider (Microsoft.ACE.OLEDB.16.0) is not available to this process.";
        }
    }
}

/// <summary>Creates empty .accdb files through ADOX so OLE DB tests run against a real database.</summary>
internal static class TestDatabase
{
    public const string Provider = "Microsoft.ACE.OLEDB.16.0";

    private static readonly Lazy<bool> Available = new(Probe);

    public static bool AceAvailable => Available.Value;

    public static string Create(string path)
    {
        var type = Type.GetTypeFromProgID("ADOX.Catalog")
            ?? throw new InvalidOperationException("ADOX.Catalog is not registered.");
        dynamic catalog = Activator.CreateInstance(type)!;
        try
        {
            dynamic connection = catalog.Create($"Provider={Provider};Data Source={path}");
            connection.Close();
            Marshal.FinalReleaseComObject(connection);
        }
        finally
        {
            Marshal.FinalReleaseComObject(catalog);
        }

        return path;
    }

    private static bool Probe()
    {
        var path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"accessmcp-probe-{Guid.NewGuid():N}.accdb");
        try
        {
            Create(path);
            return File.Exists(path);
        }
        catch (Exception)
        {
            // Not registered, wrong bitness, or the provider refused: any of these means skip.
            return false;
        }
        finally
        {
            try { File.Delete(path); } catch (IOException) { }
        }
    }
}
