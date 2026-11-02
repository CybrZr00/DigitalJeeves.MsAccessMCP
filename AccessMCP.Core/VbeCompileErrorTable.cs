using System.Reflection;
using System.Text.Json;

namespace DigitalJeeves.MsAccessMCP.Core;

/// <summary>
/// Compiled-in VBE compiler error messages. The canonical table lives in
/// ~/vbe7/VbeCompileErrors.json; scripts/sync-vbe-errors.ps1 copies it here.
/// The error number is the RT_STRING resource ID in VBE7INTL.dll.
/// </summary>
internal static class VbeCompileErrorTable
{
    private const string ResourceName = "AccessMCP.Resources.VbeCompileErrors.json";

    private static readonly Lazy<State> LazyState = new(Load, isThreadSafe: true);

    public static int Count => LazyState.Value.Messages.Count;

    public static string? SourceHash => LazyState.Value.SourceHash;

    public static bool TryGet(int? code, out string message)
    {
        if (code is int value &&
            LazyState.Value.Messages.TryGetValue(value, out var text) &&
            !string.IsNullOrWhiteSpace(text))
        {
            message = Normalize(text);
            return true;
        }

        message = string.Empty;
        return false;
    }

    public static string GetMessage(int? code, string fallback)
        => TryGet(code, out var message) ? message : fallback;

    private static string Normalize(string text)
    {
        // The resource table uses '~' for a substituted identifier; the VBE record's
        // message (when present) has the real text. Keep the template readable.
        return text.Replace("~", "<name>", StringComparison.Ordinal);
    }

    private static State Load()
    {
        try
        {
            using var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(ResourceName);
            if (stream is null)
            {
                return State.Empty;
            }

            using var document = JsonDocument.Parse(stream);
            var root = document.RootElement;
            var messages = new Dictionary<int, string>();

            if (root.TryGetProperty("messages", out var messagesElement) &&
                messagesElement.ValueKind == JsonValueKind.Object)
            {
                foreach (var property in messagesElement.EnumerateObject())
                {
                    if (int.TryParse(property.Name, out var code) &&
                        property.Value.ValueKind == JsonValueKind.String)
                    {
                        var text = property.Value.GetString();
                        if (!string.IsNullOrWhiteSpace(text))
                        {
                            messages[code] = text;
                        }
                    }
                }
            }

            var sourceHash = root.TryGetProperty("vbe7IntlSha256", out var hash) &&
                             hash.ValueKind == JsonValueKind.String
                ? hash.GetString()
                : null;

            return new State(messages, sourceHash);
        }
        catch
        {
            return State.Empty;
        }
    }

    private sealed record State(IReadOnlyDictionary<int, string> Messages, string? SourceHash)
    {
        public static State Empty { get; } = new(new Dictionary<int, string>(), null);
    }
}
