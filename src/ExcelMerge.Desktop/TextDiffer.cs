namespace ExcelMerge.Desktop;

public readonly record struct TextSpan(int Start, int Length);

/// <summary>
/// Character-level diff for cell display text. Trims the common prefix/suffix, then runs a
/// bounded LCS on the differing middle; oversized inputs fall back to a single span.
/// </summary>
public static class TextDiffer
{
    public const int MaxInputLength = 10_000;
    public const int MaxMiddleLength = 256;

    /// <summary>Returns the changed spans within <paramref name="newText"/>, ordered by start.</summary>
    public static TextSpan[] GetChangedSpans(string? oldText, string? newText)
    {
        oldText ??= string.Empty;
        newText ??= string.Empty;
        if (oldText == newText || newText.Length == 0)
        {
            return [];
        }

        if (oldText.Length == 0 ||
            oldText.Length > MaxInputLength ||
            newText.Length > MaxInputLength)
        {
            return [new TextSpan(0, newText.Length)];
        }

        var limit = Math.Min(oldText.Length, newText.Length);
        var prefix = 0;
        while (prefix < limit && oldText[prefix] == newText[prefix])
        {
            prefix++;
        }

        var suffix = 0;
        while (suffix < limit - prefix &&
               oldText[oldText.Length - 1 - suffix] == newText[newText.Length - 1 - suffix])
        {
            suffix++;
        }

        var oldMiddleLength = oldText.Length - prefix - suffix;
        var newMiddleLength = newText.Length - prefix - suffix;
        if (newMiddleLength == 0)
        {
            return [];
        }

        if (oldMiddleLength == 0 ||
            oldMiddleLength > MaxMiddleLength ||
            newMiddleLength > MaxMiddleLength)
        {
            return [new TextSpan(prefix, newMiddleLength)];
        }

        return DiffMiddle(oldText, prefix, oldMiddleLength, newText, prefix, newMiddleLength);
    }

    private static TextSpan[] DiffMiddle(
        string oldText,
        int oldStart,
        int oldLength,
        string newText,
        int newStart,
        int newLength)
    {
        var stride = newLength + 1;
        var lengths = new int[(oldLength + 1) * stride];
        for (var o = oldLength - 1; o >= 0; o--)
        {
            for (var n = newLength - 1; n >= 0; n--)
            {
                lengths[(o * stride) + n] = oldText[oldStart + o] == newText[newStart + n]
                    ? lengths[((o + 1) * stride) + n + 1] + 1
                    : Math.Max(
                        lengths[((o + 1) * stride) + n],
                        lengths[(o * stride) + n + 1]);
            }
        }

        Span<bool> matched = stackalloc bool[newLength];
        var oldIndex = 0;
        var newIndex = 0;
        while (oldIndex < oldLength && newIndex < newLength)
        {
            if (oldText[oldStart + oldIndex] == newText[newStart + newIndex])
            {
                matched[newIndex] = true;
                oldIndex++;
                newIndex++;
            }
            else if (lengths[((oldIndex + 1) * stride) + newIndex] >=
                     lengths[(oldIndex * stride) + newIndex + 1])
            {
                oldIndex++;
            }
            else
            {
                newIndex++;
            }
        }

        var spans = new List<TextSpan>();
        var runStart = -1;
        for (var index = 0; index <= newLength; index++)
        {
            var changed = index < newLength && !matched[index];
            if (changed && runStart < 0)
            {
                runStart = index;
            }
            else if (!changed && runStart >= 0)
            {
                spans.Add(new TextSpan(newStart + runStart, index - runStart));
                runStart = -1;
            }
        }

        return spans.Count == 0
            ? [new TextSpan(newStart, newLength)]
            : [.. spans];
    }
}
