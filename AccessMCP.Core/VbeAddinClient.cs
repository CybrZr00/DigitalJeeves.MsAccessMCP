using System.Globalization;
using System.Text.Json;

namespace DigitalJeeves.MsAccessMCP.Core;

/// <summary>Snapshot of the VBE compiler's BTSRC_ERR record.</summary>
internal sealed record VbeCompileRecord(
    bool Present,
    int? ErrorNumber,
    int? Line,
    int? Column,
    string? Module,
    string? Project,
    string? Message);

/// <summary>
/// Loads the AccessMCP VBE add-in on demand and reads the VBE compiler's own
/// compile-error record. The add-in is a native x86 COM component registered per-user;
/// <see cref="TryConnect"/> returns null when it is not installed or cannot be loaded,
/// so callers can fall back to Application.IsCompiled and the VBE dialog.
/// </summary>
internal sealed class VbeAddinClient : IDisposable
{
    public const string ProgId = "AccessMCP.VbeProbe";

    private readonly object _addIn;
    private readonly object _probe;
    private bool _disposed;

    private VbeAddinClient(object addIn, object probe)
    {
        _addIn = addIn;
        _probe = probe;
    }

    /// <summary>
    /// Connects the add-in through Application.VBE.AddIns and returns a client, or null
    /// when it is unavailable. Never throws.
    /// </summary>
    public static VbeAddinClient? TryConnect(object? vbe)
    {
        if (vbe is null)
        {
            return null;
        }

        object? addIns = null;
        object? addIn = null;
        object? probe = null;
        try
        {
            addIns = AccessComAutomation.GetProperty(vbe, "AddIns");
            if (addIns is null)
            {
                return null;
            }

            try { AccessComAutomation.InvokeMethod(addIns, "Update"); } catch { }

            try
            {
                addIn = AccessComAutomation.GetProperty(addIns, "Item", ProgId);
            }
            catch
            {
                return null;
            }

            if (addIn is null)
            {
                return null;
            }

            AccessComAutomation.SetProperty(addIn, "Connect", true);
            probe = AccessComAutomation.GetProperty(addIn, "Object");
            if (probe is null)
            {
                return null;
            }

            var client = new VbeAddinClient(addIn, probe);

            // Ownership transfers to the client; do not release in finally.
            addIn = null;
            probe = null;
            return client;
        }
        catch
        {
            return null;
        }
        finally
        {
            AccessComAutomation.ReleaseComObject(probe);
            if (addIn is not null)
            {
                try { AccessComAutomation.SetProperty(addIn, "Connect", false); } catch { }
                AccessComAutomation.ReleaseComObject(addIn);
            }
            AccessComAutomation.ReleaseComObject(addIns);
        }
    }

    /// <summary>
    /// Reads the current BTSRC_ERR record. Must be called on the same Access STA,
    /// immediately after the compile command, before anything calls BuildErrorInfo.
    /// </summary>
    public VbeCompileRecord ReadErrorRecord()
    {
        try
        {
            var json = Convert.ToString(
                AccessComAutomation.InvokeMethod(_probe, "ReadErrorRecord"),
                CultureInfo.InvariantCulture) ?? string.Empty;

            if (string.IsNullOrWhiteSpace(json))
            {
                return new VbeCompileRecord(false, null, null, null, null, null, null);
            }

            using var document = JsonDocument.Parse(json);
            var root = document.RootElement;
            return new VbeCompileRecord(
                GetBool(root, "present"),
                GetInt(root, "errorNumber"),
                GetInt(root, "line"),
                GetInt(root, "column"),
                GetString(root, "module"),
                GetString(root, "project"),
                GetString(root, "message"));
        }
        catch
        {
            return new VbeCompileRecord(false, null, null, null, null, null, null);
        }
    }

    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }
        _disposed = true;

        try { AccessComAutomation.SetProperty(_addIn, "Connect", false); } catch { }
        AccessComAutomation.ReleaseComObject(_probe);
        AccessComAutomation.ReleaseComObject(_addIn);
    }

    private static string? GetString(JsonElement element, string name)
        => element.TryGetProperty(name, out var value) && value.ValueKind == JsonValueKind.String
            ? value.GetString()
            : null;

    private static int? GetInt(JsonElement element, string name)
        => element.TryGetProperty(name, out var value) && value.TryGetInt32(out var parsed)
            ? parsed
            : null;

    private static bool GetBool(JsonElement element, string name)
        => element.TryGetProperty(name, out var value) && value.ValueKind == JsonValueKind.True;
}
