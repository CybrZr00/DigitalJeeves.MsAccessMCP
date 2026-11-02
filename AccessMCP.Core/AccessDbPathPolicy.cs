using System.IO;

namespace DigitalJeeves.MsAccessMCP.Core;

/// <summary>
/// Validates and normalizes Access database paths for MCP tool calls.
/// </summary>
internal static class AccessDbPathPolicy
{
    private static readonly string[] EditableExtensions = [".mdb", ".accdb"];
    private static readonly string[] CompiledExtensions = [".mde", ".accde"];

    /// <summary>
    /// Normalizes a database path, checks it exists, and validates the extension.
    /// OLE DB tools only accept editable databases.
    /// </summary>
    public static string NormalizeAndValidate(string? databasePath, AccessMcpOptions options)
    {
        var fullPath = NormalizeExisting(databasePath, options);

        var ext = Path.GetExtension(fullPath);
        if (!EditableExtensions.Contains(ext, StringComparer.OrdinalIgnoreCase))
        {
            throw new ArgumentException($"Unsupported database extension '{ext}'. Expected .mdb or .accdb.", nameof(databasePath));
        }

        return fullPath;
    }

    /// <summary>
    /// Normalizes a database path for COM-based VBA tools. Accepts editable databases and, when
    /// <paramref name="allowCompiled"/> is true, compiled .mde/.accde databases.
    /// </summary>
    public static string NormalizeComDatabasePath(string? databasePath, AccessMcpOptions options, bool allowCompiled = true)
    {
        var fullPath = NormalizeExisting(databasePath, options);

        var ext = Path.GetExtension(fullPath);
        if (!EditableExtensions.Contains(ext, StringComparer.OrdinalIgnoreCase) &&
            !(allowCompiled && CompiledExtensions.Contains(ext, StringComparer.OrdinalIgnoreCase)))
        {
            throw new ArgumentException(
                $"Unsupported database extension '{ext}'. Expected .mdb, .accdb" +
                (allowCompiled ? ", .mde or .accde." : "."),
                nameof(databasePath));
        }

        return fullPath;
    }

    /// <summary>
    /// True when the extension indicates a compiled Access database. Compiled databases cannot be
    /// edited through the VBE, but they can be opened and their procedures can run.
    /// </summary>
    public static bool IsCompiledDatabase(string path)
        => CompiledExtensions.Contains(Path.GetExtension(path), StringComparer.OrdinalIgnoreCase);

    /// <summary>
    /// A stable lowercase kind for diagnostics: mdb, accdb, mde or accde.
    /// </summary>
    public static string GetDatabaseKind(string path)
        => Path.GetExtension(path).TrimStart('.').ToLowerInvariant();

    /// <summary>
    /// Expands environment variables and normalizes a configured directory. Falls back to
    /// <paramref name="fallback"/> when nothing is configured.
    /// </summary>
    public static string NormalizeDirectory(string? configured, string fallback)
    {
        var value = string.IsNullOrWhiteSpace(configured) ? fallback : configured;
        return Path.GetFullPath(Environment.ExpandEnvironmentVariables(value));
    }

    private static string NormalizeExisting(string? databasePath, AccessMcpOptions options)
    {
        if (string.IsNullOrWhiteSpace(databasePath))
        {
            databasePath = options.DefaultDatabasePath;
        }

        if (string.IsNullOrWhiteSpace(databasePath))
        {
            throw new ArgumentException("Database path is required.", nameof(databasePath));
        }

        var fullPath = Path.GetFullPath(Environment.ExpandEnvironmentVariables(databasePath));

        if (!File.Exists(fullPath))
        {
            throw new FileNotFoundException($"Database file not found: {fullPath}", fullPath);
        }

        return fullPath;
    }
}
