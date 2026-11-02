using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace DigitalJeeves.MsAccessMCP.Core;

/// <summary>
/// Remembers a StartUpForm value that a COM session removed until it has been written back. If the
/// session dies between the remove and the restore (a watchdog kill or an Access crash during
/// <c>OpenCurrentDatabase</c>), the next session on the same database finds the record and restores
/// the value, so the database doesn't permanently lose its start-up form.
/// </summary>
internal sealed class StartUpFormRecovery(string directory)
{
    public static StartUpFormRecovery Default { get; } = new(Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "AccessMCP",
        "StartUpFormRecovery"));

    public void Save(string databasePath, string value)
    {
        Directory.CreateDirectory(directory);
        File.WriteAllLines(RecordPath(databasePath), [Path.GetFullPath(databasePath), value]);
    }

    /// <summary>Returns the recorded value for <paramref name="databasePath"/>, or null when there is none.</summary>
    public string? TryLoad(string databasePath)
    {
        try
        {
            var lines = File.ReadAllLines(RecordPath(databasePath));

            // The first line is the database path. It makes the folder readable by hand, and a record
            // that doesn't match is ignored rather than restored onto the wrong database.
            return lines.Length >= 2 && string.Equals(lines[0], Path.GetFullPath(databasePath), StringComparison.OrdinalIgnoreCase)
                ? lines[1]
                : null;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            return null;
        }
    }

    public void Clear(string databasePath)
    {
        try
        {
            File.Delete(RecordPath(databasePath));
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
        }
    }

    internal string RecordPath(string databasePath)
    {
        var key = Path.GetFullPath(databasePath).ToUpperInvariant();
        var hash = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(key)));
        return Path.Combine(directory, hash + ".txt");
    }
}
