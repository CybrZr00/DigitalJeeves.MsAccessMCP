using System.Globalization;
using System.IO;
using System.Text;
using System.Text.Json;

namespace DigitalJeeves.MsAccessMCP.Core;

/// <summary>One test outcome parsed from the kit's NDJSON file.</summary>
internal sealed record VbaTestOutcome(string Name, string Outcome, string? Message, int ElapsedMs);

/// <summary>Aggregated result of an <c>access_vba_run_tests</c> run.</summary>
internal sealed record VbaTestSummary(
    bool Ok,
    int Total,
    int Passed,
    int Failed,
    int Skipped,
    int XFailed,
    int XPassed,
    IReadOnlyList<VbaTestOutcome> Tests,
    IReadOnlyList<string> LogTail,
    string? Error,
    string? RawTail);

/// <summary>
/// Reads the line-delimited JSON written by <c>VbaTestKit</c>. Every line is flushed as it is
/// written, so a crashed or killed run still yields the results that were produced.
/// </summary>
internal static class NdjsonResults
{
    private const char Separator = '\0';
    private const string TestMarker = "@@test";

    public static VbaTestSummary ParseTests(string resultPath, int maxChars, bool markIncomplete = false)
    {
        var outcomes = new List<VbaTestOutcome>();
        var started = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        var finished = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var logs = new List<string>();
        string? error = null;
        var raw = new List<string>();
        var consumed = 0;

        foreach (var line in ReadLines(resultPath))
        {
            if (consumed >= maxChars)
            {
                break;
            }

            consumed += line.Length;
            raw.Add(line);
            if (raw.Count > 200)
            {
                raw.RemoveAt(0);
            }

            JsonDocument document;
            try
            {
                document = JsonDocument.Parse(line);
            }
            catch
            {
                // A crash can leave a partial final line. Keep what came before it.
                continue;
            }

            using (document)
            {
                var root = document.RootElement;
                var kind = GetString(root, "event");

                switch (kind)
                {
                    case "log":
                        var text = GetString(root, "text") ?? string.Empty;
                        if (text.StartsWith(TestMarker, StringComparison.Ordinal))
                        {
                            HandleTestEvent(text, outcomes, started, finished);
                        }
                        else
                        {
                            logs.Add(text);
                        }

                        break;

                    case "value":
                        logs.Add($"[value] {GetString(root, "type")}: {GetString(root, "text")}");
                        break;

                    case "error":
                        error ??= $"{GetString(root, "source")}: {GetString(root, "description")}";
                        break;
                }
            }
        }

        if (markIncomplete)
        {
            foreach (var name in started.Keys.Where(n => !finished.Contains(n)))
            {
                outcomes.Add(new VbaTestOutcome(name, "error", "No result was written (the run crashed or timed out).", 0));
                finished.Add(name);
            }
        }

        var ordered = outcomes
            .OrderBy(o => o.Name, StringComparer.OrdinalIgnoreCase)
            .ToArray();

        var passed = ordered.Count(o => o.Outcome == "pass");
        var failed = ordered.Count(o => o.Outcome is "fail" or "error" or "xpass");
        var skipped = ordered.Count(o => o.Outcome == "skip");
        var xfailed = ordered.Count(o => o.Outcome == "xfail");
        var xpassed = ordered.Count(o => o.Outcome == "xpass");

        var ok = failed == 0 && error is null && outcomes.All(o => o.Outcome != "error");

        return new VbaTestSummary(
            ok,
            ordered.Length,
            passed,
            failed,
            skipped,
            xfailed,
            xpassed,
            ordered,
            CapLog(logs, maxChars),
            error,
            string.Join("\n", raw));
    }

    /// <summary>Reads the kit's plain log lines (no test events) for a non-test run.</summary>
    public static (IReadOnlyList<string> Log, string? Error) ParseLog(string resultPath, int maxChars)
    {
        var logs = new List<string>();
        string? error = null;

        foreach (var line in ReadLines(resultPath))
        {
            try
            {
                using var document = JsonDocument.Parse(line);
                var root = document.RootElement;
                switch (GetString(root, "event"))
                {
                    case "log":
                        var text = GetString(root, "text") ?? string.Empty;
                        if (!text.StartsWith(TestMarker, StringComparison.Ordinal))
                        {
                            logs.Add(text);
                        }

                        break;
                    case "value":
                        logs.Add($"[value] {GetString(root, "type")}: {GetString(root, "text")}");
                        break;
                    case "error":
                        error = $"{GetString(root, "source")}: {GetString(root, "description")}";
                        break;
                }
            }
            catch
            {
                // Ignore a partial final line.
            }
        }

        return (CapLog(logs, maxChars), error);
    }

    private static void HandleTestEvent(
        string text,
        List<VbaTestOutcome> outcomes,
        Dictionary<string, int> started,
        HashSet<string> finished)
    {
        var parts = text.Split(Separator, 5);
        if (parts.Length < 3)
        {
            return;
        }

        var outcome = parts[1];
        var name = parts[2];

        // Newer runners append a tick count and a detail; older ones only a detail.
        var tick = 0;
        var detailIndex = 3;
        if (parts.Length >= 5 && int.TryParse(parts[3], NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed))
        {
            tick = parsed;
            detailIndex = 4;
        }

        var detail = parts.Length > detailIndex ? parts[detailIndex] : null;

        if (outcome.Equals("start", StringComparison.OrdinalIgnoreCase))
        {
            started[name] = tick;
            return;
        }

        started.TryGetValue(name, out var startTick);
        var elapsed = tick > 0 && startTick > 0 ? Math.Max(0, tick - startTick) : 0;

        outcomes.RemoveAll(o => o.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
        outcomes.Add(new VbaTestOutcome(name, outcome.ToLowerInvariant(), NullIfEmpty(detail), elapsed));
        finished.Add(name);
    }

    private static IReadOnlyList<string> CapLog(List<string> logs, int maxChars)
    {
        var selected = new List<string>();
        var total = 0;
        for (var i = logs.Count - 1; i >= 0; i--)
        {
            total += logs[i].Length + 1;
            if (total > maxChars)
            {
                break;
            }

            selected.Insert(0, logs[i]);
        }

        return selected;
    }

    private static IEnumerable<string> ReadLines(string path)
    {
        if (!File.Exists(path))
        {
            yield break;
        }

        using var reader = new StreamReader(path, Encoding.UTF8, detectEncodingFromByteOrderMarks: true);
        while (reader.ReadLine() is { } line)
        {
            if (line.Length > 0)
            {
                yield return line;
            }
        }
    }

    private static string? GetString(JsonElement element, string property)
        => element.TryGetProperty(property, out var value) && value.ValueKind == JsonValueKind.String
            ? value.GetString()
            : null;

    private static string? NullIfEmpty(string? value) => string.IsNullOrWhiteSpace(value) ? null : value;
}
