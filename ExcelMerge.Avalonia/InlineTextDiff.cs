#nullable enable

using Avalonia.Controls;
using Avalonia.Controls.Documents;
using Avalonia.Media;
using NetDiff;
using System.Text;

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

        var localSegments = new SegmentBuilder();
        var remoteSegments = new SegmentBuilder();
        var lineDiff = Optimize(SplitLines(local), SplitLines(remote));
        foreach (var line in lineDiff)
        {
            switch (line.Status)
            {
                case DiffStatus.Equal:
                    localSegments.Append(line.Obj1, false);
                    remoteSegments.Append(line.Obj2, false);
                    break;
                case DiffStatus.Modified:
                    AppendCharacterDiff(line.Obj1, line.Obj2, localSegments, remoteSegments);
                    break;
                case DiffStatus.Deleted:
                    localSegments.Append(line.Obj1, true);
                    break;
                case DiffStatus.Inserted:
                    remoteSegments.Append(line.Obj2, true);
                    break;
            }
        }

        return new InlineDiffResult(localSegments.Build(), remoteSegments.Build());
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
        SegmentBuilder localSegments,
        SegmentBuilder remoteSegments)
    {
        foreach (var character in Optimize(local, remote))
        {
            switch (character.Status)
            {
                case DiffStatus.Equal:
                    localSegments.Append(character.Obj1, false);
                    remoteSegments.Append(character.Obj2, false);
                    break;
                case DiffStatus.Modified:
                    localSegments.Append(character.Obj1, true);
                    remoteSegments.Append(character.Obj2, true);
                    break;
                case DiffStatus.Deleted:
                    localSegments.Append(character.Obj1, true);
                    break;
                case DiffStatus.Inserted:
                    remoteSegments.Append(character.Obj2, true);
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

    private sealed class SegmentBuilder
    {
        private readonly List<InlineDiffSegment> _segments = new();
        private readonly StringBuilder _buffer = new();
        private bool? _changed;

        public void Append(string text, bool changed)
        {
            if (string.IsNullOrEmpty(text))
                return;
            StartSegment(changed);
            _buffer.Append(text);
        }

        public void Append(char value, bool changed)
        {
            StartSegment(changed);
            _buffer.Append(value);
        }

        public IReadOnlyList<InlineDiffSegment> Build()
        {
            Flush();
            return _segments.Count == 0 ? Array.Empty<InlineDiffSegment>() : _segments;
        }

        private void StartSegment(bool changed)
        {
            if (_changed.HasValue && _changed.Value != changed)
                Flush();
            _changed = changed;
        }

        private void Flush()
        {
            if (_buffer.Length == 0 || !_changed.HasValue)
                return;
            _segments.Add(new InlineDiffSegment(_buffer.ToString(), _changed.Value));
            _buffer.Clear();
        }
    }
}
