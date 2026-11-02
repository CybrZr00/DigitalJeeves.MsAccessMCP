using System.Globalization;
using System.Text;

namespace DigitalJeeves.MsAccessMCP.Core;

/// <summary>A VBA module's code as read from one database.</summary>
internal sealed record VbaModuleSnapshot(string Name, int ComponentType, string Code);

/// <summary>A saved query's SQL as read from one database.</summary>
internal sealed record QuerySnapshot(string Name, int QueryType, string Sql);

/// <summary>A module that exists in only one of the two databases compared.</summary>
internal sealed record VbaModuleOnly(string Name, int ComponentType, int LineCount);

/// <summary>A module that exists in both databases but whose code differs.</summary>
internal sealed record VbaModuleChange(
    string Name,
    int ComponentType,
    int SourceLineCount,
    int TargetLineCount,
    string Diff);

/// <summary>A query that exists in only one of the two databases compared.</summary>
internal sealed record QueryOnly(string Name, int QueryType, int LineCount);

/// <summary>A query that exists in both databases but whose SQL differs.</summary>
internal sealed record QueryChange(
    string Name,
    int QueryType,
    int SourceLineCount,
    int TargetLineCount,
    string Diff);

/// <summary>Every module and (optionally) query of one database, plus the VBA project they came from.</summary>
internal sealed record DbSnapshotSet(
    string ProjectName,
    IReadOnlyList<VbaModuleSnapshot> Modules,
    IReadOnlyList<QuerySnapshot> Queries);

/// <summary>The result of comparing the VBA modules of two databases.</summary>
internal sealed record VbaModuleDiffResult(
    IReadOnlyList<VbaModuleOnly> OnlyInSource,
    IReadOnlyList<VbaModuleOnly> OnlyInTarget,
    IReadOnlyList<VbaModuleChange> Changed,
    IReadOnlyList<VbaModuleOnly> Unchanged);

/// <summary>The result of comparing the saved queries of two databases.</summary>
internal sealed record QueryDiffResult(
    IReadOnlyList<QueryOnly> OnlyInSource,
    IReadOnlyList<QueryOnly> OnlyInTarget,
    IReadOnlyList<QueryChange> Changed,
    IReadOnlyList<QueryOnly> Unchanged);

/// <summary>
/// Pure comparison logic for the VBA modules of two Access databases. Deliberately free of COM/OLE so the
/// algorithm can be unit-tested without Access installed - <see cref="AccessTools.DiffVbaModulesAsync"/>
/// reads the snapshots and hands them here.
/// </summary>
internal static class VbaModuleDiff
{
    /// <summary>
    /// Classifies each module by name as only-in-source, only-in-target, changed (with a unified diff) or
    /// unchanged.
    /// </summary>
    public static VbaModuleDiffResult Compare(
        IReadOnlyList<VbaModuleSnapshot> source,
        IReadOnlyList<VbaModuleSnapshot> target,
        int contextLines,
        bool ignoreCase = true,
        bool ignoreWhitespace = true,
        string sourceLabel = "source",
        string targetLabel = "target")
    {
        ArgumentNullException.ThrowIfNull(source);
        ArgumentNullException.ThrowIfNull(target);

        var result = NamedTextDiff.Compare(
            source.Select(module => new NamedTextSnapshot(module.Name, module.ComponentType, module.Code)).ToList(),
            target.Select(module => new NamedTextSnapshot(module.Name, module.ComponentType, module.Code)).ToList(),
            contextLines,
            ignoreCase,
            ignoreWhitespace,
            sourceLabel,
            targetLabel);

        return new VbaModuleDiffResult(
            result.OnlyInSource.Select(item => new VbaModuleOnly(item.Name, item.Type, item.LineCount)).ToList(),
            result.OnlyInTarget.Select(item => new VbaModuleOnly(item.Name, item.Type, item.LineCount)).ToList(),
            result.Changed.Select(item => new VbaModuleChange(item.Name, item.Type, item.SourceLineCount, item.TargetLineCount, item.Diff)).ToList(),
            result.Unchanged.Select(item => new VbaModuleOnly(item.Name, item.Type, item.LineCount)).ToList());
    }

    /// <summary>Number of lines in a module body, treating an empty body as zero lines.</summary>
    public static int CountLines(string code)
    {
        if (string.IsNullOrEmpty(code))
        {
            return 0;
        }

        var count = 1;
        foreach (var c in code)
        {
            if (c == '\n')
            {
                count++;
            }
        }

        return count;
    }
}

/// <summary>
/// Pure comparison logic for the saved queries of two Access databases, sharing the classification and
/// unified-diff rules with the module comparison.
/// </summary>
internal static class QueryDiff
{
    /// <summary>
    /// Classifies each query by name as only-in-source, only-in-target, changed (with a unified diff) or
    /// unchanged.
    /// </summary>
    public static QueryDiffResult Compare(
        IReadOnlyList<QuerySnapshot> source,
        IReadOnlyList<QuerySnapshot> target,
        int contextLines,
        bool ignoreCase = true,
        bool ignoreWhitespace = true,
        string sourceLabel = "source",
        string targetLabel = "target")
    {
        ArgumentNullException.ThrowIfNull(source);
        ArgumentNullException.ThrowIfNull(target);

        var result = NamedTextDiff.Compare(
            source.Select(query => new NamedTextSnapshot(query.Name, query.QueryType, query.Sql)).ToList(),
            target.Select(query => new NamedTextSnapshot(query.Name, query.QueryType, query.Sql)).ToList(),
            contextLines,
            ignoreCase,
            ignoreWhitespace,
            sourceLabel,
            targetLabel);

        return new QueryDiffResult(
            result.OnlyInSource.Select(item => new QueryOnly(item.Name, item.Type, item.LineCount)).ToList(),
            result.OnlyInTarget.Select(item => new QueryOnly(item.Name, item.Type, item.LineCount)).ToList(),
            result.Changed.Select(item => new QueryChange(item.Name, item.Type, item.SourceLineCount, item.TargetLineCount, item.Diff)).ToList(),
            result.Unchanged.Select(item => new QueryOnly(item.Name, item.Type, item.LineCount)).ToList());
    }
}

/// <summary>A named piece of text to compare - a VBA module body or a saved query's SQL.</summary>
internal sealed record NamedTextSnapshot(string Name, int Type, string Text);

internal sealed record NamedTextOnly(string Name, int Type, int LineCount);

internal sealed record NamedTextChange(
    string Name,
    int Type,
    int SourceLineCount,
    int TargetLineCount,
    string Diff);

internal sealed record NamedTextDiffResult(
    IReadOnlyList<NamedTextOnly> OnlyInSource,
    IReadOnlyList<NamedTextOnly> OnlyInTarget,
    IReadOnlyList<NamedTextChange> Changed,
    IReadOnlyList<NamedTextOnly> Unchanged);

/// <summary>
/// The classification and unified-diff loop shared by the module and query comparisons, so the
/// ignoreCase/ignoreWhitespace rules stay identical for both.
/// </summary>
internal static class NamedTextDiff
{
    public static NamedTextDiffResult Compare(
        IReadOnlyList<NamedTextSnapshot> source,
        IReadOnlyList<NamedTextSnapshot> target,
        int contextLines,
        bool ignoreCase,
        bool ignoreWhitespace,
        string sourceLabel,
        string targetLabel)
    {
        ArgumentNullException.ThrowIfNull(source);
        ArgumentNullException.ThrowIfNull(target);
        if (contextLines < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(contextLines), "Context lines cannot be negative.");
        }

        var sourceByName = ToDictionary(source);
        var targetByName = ToDictionary(target);

        var onlyInSource = new List<NamedTextOnly>();
        var onlyInTarget = new List<NamedTextOnly>();
        var changed = new List<NamedTextChange>();
        var unchanged = new List<NamedTextOnly>();

        foreach (var item in source)
        {
            if (!targetByName.TryGetValue(item.Name, out var other))
            {
                onlyInSource.Add(new NamedTextOnly(item.Name, item.Type, VbaModuleDiff.CountLines(item.Text)));
                continue;
            }

            if (UnifiedDiff.AreEquivalent(item.Text, other.Text, ignoreCase, ignoreWhitespace))
            {
                unchanged.Add(new NamedTextOnly(item.Name, item.Type, VbaModuleDiff.CountLines(item.Text)));
                continue;
            }

            var diff = UnifiedDiff.Build(
                item.Text,
                other.Text,
                $"{sourceLabel}/{item.Name}",
                $"{targetLabel}/{item.Name}",
                contextLines,
                ignoreCase,
                ignoreWhitespace);

            changed.Add(new NamedTextChange(
                item.Name,
                item.Type,
                VbaModuleDiff.CountLines(item.Text),
                VbaModuleDiff.CountLines(other.Text),
                diff));
        }

        foreach (var item in target)
        {
            if (!sourceByName.ContainsKey(item.Name))
            {
                onlyInTarget.Add(new NamedTextOnly(item.Name, item.Type, VbaModuleDiff.CountLines(item.Text)));
            }
        }

        return new NamedTextDiffResult(onlyInSource, onlyInTarget, changed, unchanged);
    }

    private static Dictionary<string, NamedTextSnapshot> ToDictionary(IReadOnlyList<NamedTextSnapshot> items)
    {
        var map = new Dictionary<string, NamedTextSnapshot>(StringComparer.OrdinalIgnoreCase);
        foreach (var item in items)
        {
            // First wins; module and query names are unique within a database anyway.
            map.TryAdd(item.Name, item);
        }

        return map;
    }
}

/// <summary>
/// Minimal line-oriented unified diff, used to render VBA module changes. Self-contained (no external
/// dependency) and memory-bounded so a pair of large, wildly different modules can't exhaust the process:
/// instead of Myers' O((N+M)^2) trace, it uses Hirschberg's linear-space LCS, and downgrades anything past
/// <see cref="MaxLcsCells"/> to a wholesale replacement.
/// </summary>
internal static class UnifiedDiff
{
    private enum LineOp
    {
        Equal,
        Delete,
        Insert
    }

    private readonly record struct Edit(LineOp Op, string Text);

    private readonly record struct DiffLine(LineOp Op, string Text, int SourceLine, int TargetLine);

    /// <summary>
    /// Above this many cells (source lines x target lines) the fine-grained LCS is skipped and the changed
    /// middle is reported as an all-delete/all-insert block. VBA modules are far below this in practice.
    /// </summary>
    private const long MaxLcsCells = 25_000_000;

    public static string Build(
        string sourceText,
        string targetText,
        string sourceLabel,
        string targetLabel,
        int contextLines,
        bool ignoreCase = true,
        bool ignoreWhitespace = true)
    {
        if (contextLines < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(contextLines), "Context lines cannot be negative.");
        }

        var source = SplitLines(sourceText);
        var target = SplitLines(targetText);
        var comparer = ignoreCase ? StringComparer.OrdinalIgnoreCase : StringComparer.Ordinal;

        // Comparison keys are computed once per line: the LCS loops below compare O(N*M) times, so
        // normalizing inside the comparison would allocate millions of strings for a large module.
        var sourceKeys = BuildKeys(source, ignoreWhitespace);
        var targetKeys = BuildKeys(target, ignoreWhitespace);

        var builder = new StringBuilder();
        builder.Append("--- ").Append(sourceLabel).Append('\n');
        builder.Append("+++ ").Append(targetLabel).Append('\n');

        AppendHunks(builder, BuildDiffLines(source, sourceKeys, target, targetKeys, comparer), contextLines);
        return builder.ToString();
    }

    /// <summary>
    /// True when two texts are the same after the requested case/whitespace rules, compared line by line
    /// so CRLF and LF never count as a difference on their own.
    /// </summary>
    public static bool AreEquivalent(string sourceText, string targetText, bool ignoreCase, bool ignoreWhitespace)
    {
        var source = SplitLines(sourceText);
        var target = SplitLines(targetText);
        if (source.Length != target.Length)
        {
            return false;
        }

        var comparer = ignoreCase ? StringComparer.OrdinalIgnoreCase : StringComparer.Ordinal;
        for (var i = 0; i < source.Length; i++)
        {
            var left = ignoreWhitespace ? CollapseWhitespace(source[i]) : source[i];
            var right = ignoreWhitespace ? CollapseWhitespace(target[i]) : target[i];
            if (!comparer.Equals(left, right))
            {
                return false;
            }
        }

        return true;
    }

    /// <summary>
    /// The whitespace-insensitive form of one line: ends trimmed and every internal run of whitespace
    /// collapsed to a single space (git's --ignore-space-change). Whitespace added or removed where there
    /// was none - "x=1" vs "x = 1" - still reads as a difference.
    /// </summary>
    internal static string CollapseWhitespace(string line)
    {
        var start = 0;
        var end = line.Length;
        while (start < end && char.IsWhiteSpace(line[start]))
        {
            start++;
        }

        while (end > start && char.IsWhiteSpace(line[end - 1]))
        {
            end--;
        }

        if (start == end)
        {
            return string.Empty;
        }

        var builder = new StringBuilder(end - start);
        var inWhitespace = false;
        for (var i = start; i < end; i++)
        {
            var c = line[i];
            if (char.IsWhiteSpace(c))
            {
                if (!inWhitespace)
                {
                    builder.Append(' ');
                    inWhitespace = true;
                }
            }
            else
            {
                builder.Append(c);
                inWhitespace = false;
            }
        }

        return builder.ToString();
    }

    private static string[] BuildKeys(string[] lines, bool ignoreWhitespace)
    {
        if (!ignoreWhitespace)
        {
            return lines;
        }

        var keys = new string[lines.Length];
        for (var i = 0; i < lines.Length; i++)
        {
            keys[i] = CollapseWhitespace(lines[i]);
        }

        return keys;
    }

    private static string[] SplitLines(string text)
        => string.IsNullOrEmpty(text)
            ? []
            : text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');

    /// <summary>
    /// Produces the full equal/delete/insert line script with 1-based source and target line numbers
    /// attached. An insert line's <see cref="DiffLine.SourceLine"/> is the source line it precedes.
    /// </summary>
    private static List<DiffLine> BuildDiffLines(
        string[] source,
        string[] sourceKeys,
        string[] target,
        string[] targetKeys,
        StringComparer comparer)
    {
        // Common prefix/suffix is equal by definition; trimming it keeps the LCS search space small for
        // the usual "a handful of lines changed" case, which is what VBA edits look like.
        var prefix = 0;
        while (prefix < source.Length && prefix < target.Length && comparer.Equals(sourceKeys[prefix], targetKeys[prefix]))
        {
            prefix++;
        }

        var suffix = 0;
        while (suffix < source.Length - prefix
               && suffix < target.Length - prefix
               && comparer.Equals(sourceKeys[source.Length - 1 - suffix], targetKeys[target.Length - 1 - suffix]))
        {
            suffix++;
        }

        var sourceMiddle = source.AsSpan(prefix, source.Length - prefix - suffix).ToArray();
        var sourceMiddleKeys = sourceKeys.AsSpan(prefix, sourceKeys.Length - prefix - suffix).ToArray();
        var targetMiddle = target.AsSpan(prefix, target.Length - prefix - suffix).ToArray();
        var targetMiddleKeys = targetKeys.AsSpan(prefix, targetKeys.Length - prefix - suffix).ToArray();

        var edits = new List<Edit>(source.Length + target.Length);
        for (var i = 0; i < prefix; i++)
        {
            edits.Add(new Edit(LineOp.Equal, source[i]));
        }

        AppendMiddle(edits, sourceMiddle, sourceMiddleKeys, targetMiddle, targetMiddleKeys, comparer);

        for (var i = 0; i < suffix; i++)
        {
            edits.Add(new Edit(LineOp.Equal, source[prefix + sourceMiddle.Length + i]));
        }

        var diffLines = new List<DiffLine>(edits.Count);
        var sourceLine = 1;
        var targetLine = 1;
        foreach (var edit in edits)
        {
            diffLines.Add(new DiffLine(edit.Op, edit.Text, sourceLine, targetLine));
            if (edit.Op != LineOp.Insert)
            {
                sourceLine++;
            }

            if (edit.Op != LineOp.Delete)
            {
                targetLine++;
            }
        }

        return diffLines;
    }

    private static void AppendMiddle(
        List<Edit> edits,
        string[] source,
        string[] sourceKeys,
        string[] target,
        string[] targetKeys,
        StringComparer comparer)
    {
        if ((long)source.Length * target.Length > MaxLcsCells)
        {
            foreach (var line in source)
            {
                edits.Add(new Edit(LineOp.Delete, line));
            }

            foreach (var line in target)
            {
                edits.Add(new Edit(LineOp.Insert, line));
            }

            return;
        }

        var matched = LcsPairs(sourceKeys, targetKeys, comparer);
        int i = 0, j = 0, p = 0;
        while (i < source.Length && j < target.Length)
        {
            if (p < matched.Count && matched[p].Source == i && matched[p].Target == j)
            {
                edits.Add(new Edit(LineOp.Equal, source[i]));
                i++;
                j++;
                p++;
            }
            else if (p < matched.Count && i < matched[p].Source)
            {
                edits.Add(new Edit(LineOp.Delete, source[i]));
                i++;
            }
            else if (p < matched.Count)
            {
                edits.Add(new Edit(LineOp.Insert, target[j]));
                j++;
            }
            else
            {
                // No further matches: drain the remaining source as deletions before inserting, so a
                // replaced block reads as -old then +new (the usual unified-diff ordering).
                edits.Add(new Edit(LineOp.Delete, source[i]));
                i++;
            }
        }

        while (i < source.Length)
        {
            edits.Add(new Edit(LineOp.Delete, source[i++]));
        }

        while (j < target.Length)
        {
            edits.Add(new Edit(LineOp.Insert, target[j++]));
        }
    }

    /// <summary>
    /// Hirschberg's algorithm: the matched (equal) line index pairs of the longest common subsequence, in
    /// increasing order, using O(min(N,M)) working memory per recursion level.
    /// </summary>
    private static List<(int Source, int Target)> LcsPairs(string[] source, string[] target, StringComparer comparer)
    {
        var pairs = new List<(int Source, int Target)>();
        LcsPairsCore(source, 0, source.Length, target, 0, target.Length, pairs, comparer);
        return pairs;
    }

    private static void LcsPairsCore(
        string[] a,
        int aLo,
        int aHi,
        string[] b,
        int bLo,
        int bHi,
        List<(int Source, int Target)> pairs,
        StringComparer comparer)
    {
        if (aLo >= aHi || bLo >= bHi)
        {
            return;
        }

        // Single source line: find the first matching target line, or report no match.
        if (aHi - aLo == 1)
        {
            for (var j = bLo; j < bHi; j++)
            {
                if (comparer.Equals(a[aLo], b[j]))
                {
                    pairs.Add((aLo, j));
                    return;
                }
            }

            return;
        }

        var mid = aLo + ((aHi - aLo) / 2);
        var prefixLengths = LcsLengthsPrefix(a, aLo, mid, b, bLo, bHi, comparer);
        var suffixLengths = LcsLengthsSuffix(a, mid, aHi, b, bLo, bHi, comparer);

        // Split b where the two halves' LCS lengths sum highest: left half a[aLo..mid) with b[bLo..k),
        // right half a[mid..aHi) with b[k..bHi).
        var bestK = 0;
        var best = -1;
        for (var k = 0; k <= bHi - bLo; k++)
        {
            var value = prefixLengths[k] + suffixLengths[k];
            if (value > best)
            {
                best = value;
                bestK = k;
            }
        }

        LcsPairsCore(a, aLo, mid, b, bLo, bLo + bestK, pairs, comparer);
        LcsPairsCore(a, mid, aHi, b, bLo + bestK, bHi, pairs, comparer);
    }

    /// <summary>result[k] = LCS length of a[aLo..aHi) against b[bLo..bLo+k).</summary>
    private static int[] LcsLengthsPrefix(string[] a, int aLo, int aHi, string[] b, int bLo, int bHi, StringComparer comparer)
    {
        var bLen = bHi - bLo;
        var previous = new int[bLen + 1];
        for (var i = aLo; i < aHi; i++)
        {
            var current = new int[bLen + 1];
            for (var k = 1; k <= bLen; k++)
            {
                current[k] = comparer.Equals(a[i], b[bLo + k - 1])
                    ? previous[k - 1] + 1
                    : Math.Max(previous[k], current[k - 1]);
            }

            previous = current;
        }

        return previous;
    }

    /// <summary>result[k] = LCS length of a[aLo..aHi) against b[bLo+k..bHi).</summary>
    private static int[] LcsLengthsSuffix(string[] a, int aLo, int aHi, string[] b, int bLo, int bHi, StringComparer comparer)
    {
        var bLen = bHi - bLo;
        var previous = new int[bLen + 1];
        for (var i = aLo; i < aHi; i++)
        {
            var current = new int[bLen + 1];
            for (var k = bLen - 1; k >= 0; k--)
            {
                current[k] = comparer.Equals(a[i], b[bLo + k])
                    ? previous[k + 1] + 1
                    : Math.Max(previous[k], current[k + 1]);
            }

            previous = current;
        }

        return previous;
    }

    private static void AppendHunks(StringBuilder builder, IReadOnlyList<DiffLine> lines, int contextLines)
    {
        // Runs of changed lines, merged when separated by no more than 2*context equal lines (so their
        // context windows would otherwise touch or overlap).
        var ranges = new List<(int Start, int End)>();
        for (var i = 0; i < lines.Count; i++)
        {
            if (lines[i].Op == LineOp.Equal)
            {
                continue;
            }

            if (ranges.Count > 0 && i - ranges[^1].End - 1 <= 2 * contextLines)
            {
                ranges[^1] = (ranges[^1].Start, i);
            }
            else
            {
                ranges.Add((i, i));
            }
        }

        foreach (var (start, end) in ranges)
        {
            var hunkStart = Math.Max(0, start - contextLines);
            var hunkEnd = Math.Min(lines.Count - 1, end + contextLines);

            var sourceCount = 0;
            var targetCount = 0;
            for (var i = hunkStart; i <= hunkEnd; i++)
            {
                if (lines[i].Op != LineOp.Insert)
                {
                    sourceCount++;
                }

                if (lines[i].Op != LineOp.Delete)
                {
                    targetCount++;
                }
            }

            builder.Append("@@ -")
                .Append(FormatRange(lines[hunkStart].SourceLine, sourceCount))
                .Append(" +")
                .Append(FormatRange(lines[hunkStart].TargetLine, targetCount))
                .Append(" @@\n");

            for (var i = hunkStart; i <= hunkEnd; i++)
            {
                var line = lines[i];
                builder.Append(line.Op switch
                {
                    LineOp.Equal => ' ',
                    LineOp.Delete => '-',
                    _ => '+'
                });
                builder.Append(line.Text);
                builder.Append('\n');
            }
        }
    }

    private static string FormatRange(int start, int count)
    {
        // Unified diff convention: an empty range points at the line *before* the insertion point, and a
        // single-line range omits the length.
        var displayStart = count == 0 ? start - 1 : start;
        return count == 1
            ? displayStart.ToString(CultureInfo.InvariantCulture)
            : string.Create(CultureInfo.InvariantCulture, $"{displayStart},{count}");
    }
}
