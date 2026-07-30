#nullable enable

using Avalonia.Controls;
using Avalonia.Controls.Documents;
using Avalonia.Media;
using NetDiff;

namespace ExcelMerge.Avalonia;

internal readonly record struct InlineDiffSegment(string Text, bool IsChanged);

internal sealed record InlineDiffResult(
    IReadOnlyList<InlineDiffSegment> Local,
    IReadOnlyList<InlineDiffSegment> Remote);

internal static class InlineTextDiff
{
    private const int MaximumDetailedLength = 20000;
    private static readonly IBrush ChangedBrush = new SolidColorBrush(Color.Parse("#FFE28A"));

    public static InlineDiffResult Create(string local, string remote)
    {
        if (string.Equals(local, remote, StringComparison.Ordinal))
        {
            var equal = string.IsNullOrEmpty(local)
                ? Array.Empty<InlineDiffSegment>()
                : new[] { new InlineDiffSegment(local, false) };
            return new InlineDiffResult(equal, equal);
        }

        if (local.Length + remote.Length > MaximumDetailedLength)
        {
            return new InlineDiffResult(
                Segments(local, true),
                Segments(remote, true));
        }

        var localSegments = new List<InlineDiffSegment>();
        var remoteSegments = new List<InlineDiffSegment>();
        var lineDiff = Optimize(SplitLines(local), SplitLines(remote));
        foreach (var line in lineDiff)
        {
            switch (line.Status)
            {
                case DiffStatus.Equal:
                    Append(localSegments, line.Obj1, false);
                    Append(remoteSegments, line.Obj2, false);
                    break;
                case DiffStatus.Modified:
                    AppendCharacterDiff(line.Obj1, line.Obj2, localSegments, remoteSegments);
                    break;
                case DiffStatus.Deleted:
                    Append(localSegments, line.Obj1, true);
                    break;
                case DiffStatus.Inserted:
                    Append(remoteSegments, line.Obj2, true);
                    break;
            }
        }

        return new InlineDiffResult(localSegments, remoteSegments);
    }

    public static void Apply(SelectableTextBlock target, IReadOnlyList<InlineDiffSegment> segments)
    {
        target.Inlines!.Clear();
        foreach (var segment in segments)
        {
            target.Inlines.Add(new Run(segment.Text)
            {
                Background = segment.IsChanged ? ChangedBrush : Brushes.Transparent,
            });
        }
    }

    private static void AppendCharacterDiff(
        string local,
        string remote,
        ICollection<InlineDiffSegment> localSegments,
        ICollection<InlineDiffSegment> remoteSegments)
    {
        foreach (var character in Optimize(local, remote))
        {
            switch (character.Status)
            {
                case DiffStatus.Equal:
                    Append(localSegments, character.Obj1.ToString(), false);
                    Append(remoteSegments, character.Obj2.ToString(), false);
                    break;
                case DiffStatus.Modified:
                    Append(localSegments, character.Obj1.ToString(), true);
                    Append(remoteSegments, character.Obj2.ToString(), true);
                    break;
                case DiffStatus.Deleted:
                    Append(localSegments, character.Obj1.ToString(), true);
                    break;
                case DiffStatus.Inserted:
                    Append(remoteSegments, character.Obj2.ToString(), true);
                    break;
            }
        }
    }

    private static IReadOnlyList<DiffResult<T>> Optimize<T>(IEnumerable<T> local, IEnumerable<T> remote)
    {
        var diff = DiffUtil.Diff(local, remote);
        diff = DiffUtil.Order(diff, DiffOrderType.LazyDeleteFirst);
        return DiffUtil.OptimizeCaseDeletedFirst(diff).ToList();
    }

    private static IReadOnlyList<string> SplitLines(string value)
    {
        if (value.Length == 0)
            return Array.Empty<string>();

        var lines = new List<string>();
        var start = 0;
        for (var index = 0; index < value.Length; index++)
        {
            if (value[index] != '\n')
                continue;
            lines.Add(value[start..(index + 1)]);
            start = index + 1;
        }
        if (start < value.Length)
            lines.Add(value[start..]);
        return lines;
    }

    private static IReadOnlyList<InlineDiffSegment> Segments(string value, bool changed) =>
        string.IsNullOrEmpty(value)
            ? Array.Empty<InlineDiffSegment>()
            : new[] { new InlineDiffSegment(value, changed) };

    private static void Append(ICollection<InlineDiffSegment> segments, string text, bool changed)
    {
        if (string.IsNullOrEmpty(text))
            return;

        if (segments is List<InlineDiffSegment> list
            && list.Count > 0
            && list[^1].IsChanged == changed)
        {
            list[^1] = list[^1] with { Text = list[^1].Text + text };
            return;
        }
        segments.Add(new InlineDiffSegment(text, changed));
    }
}
