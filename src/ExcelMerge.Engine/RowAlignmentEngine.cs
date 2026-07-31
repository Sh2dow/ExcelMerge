using ExcelMerge.Domain;

namespace ExcelMerge.Engine;

public sealed class RowAlignmentEngine
{
    public async ValueTask<RowAlignmentResult> AlignAsync(
        IWorksheetSnapshot baseSheet,
        IWorksheetSnapshot sourceSheet,
        WorksheetComparisonOptions? options = null,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(baseSheet);
        ArgumentNullException.ThrowIfNull(sourceSheet);

        var normalized = new NormalizedComparisonOptions(options ?? new WorksheetComparisonOptions());
        var context = new ComparisonContext(normalized);
        var materializedBase = await WorksheetMaterializer.ReadAsync(
            baseSheet,
            normalized,
            cancellationToken).ConfigureAwait(false);
        var materializedSource = await WorksheetMaterializer.ReadAsync(
            sourceSheet,
            normalized,
            cancellationToken).ConfigureAwait(false);

        var preparedBase = PreparedRows.Create(materializedBase.Rows, context, cancellationToken);
        var preparedSource = PreparedRows.Create(materializedSource.Rows, context, cancellationToken);
        var steps = RowAligner.Align(preparedBase, preparedSource, context, cancellationToken);
        var result = new RowAlignmentEntry[steps.Length];
        for (var index = 0; index < steps.Length; index++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var step = steps[index];
            var baseRowIndex = step.BaseOrdinal >= 0
                ? preparedBase.Rows[step.BaseOrdinal].RowIndex
                : (int?)null;
            var sourceRowIndex = step.SourceOrdinal >= 0
                ? preparedSource.Rows[step.SourceOrdinal].RowIndex
                : (int?)null;
            var kind = RowAligner.GetChangeKind(step, preparedBase, preparedSource, context, cancellationToken);
            result[index] = new RowAlignmentEntry(baseRowIndex, sourceRowIndex, kind);
        }

        return new RowAlignmentResult(
            materializedBase.Metadata,
            materializedSource.Metadata,
            result);
    }
}

internal sealed class PreparedRows
{
    private PreparedRows(
        RowRecord[] rows,
        RowSignature[] signatures,
        RowKeySignature[] keySignatures)
    {
        Rows = rows;
        Signatures = signatures;
        KeySignatures = keySignatures;
    }

    public RowRecord[] Rows { get; }

    public RowSignature[] Signatures { get; }

    public RowKeySignature[] KeySignatures { get; }

    public static PreparedRows Create(
        RowRecord[] rows,
        ComparisonContext context,
        CancellationToken cancellationToken)
    {
        var signatures = new RowSignature[rows.Length];
        var keys = context.Options.KeyColumns.Length == 0
            ? Array.Empty<RowKeySignature>()
            : new RowKeySignature[rows.Length];

        for (var index = 0; index < rows.Length; index++)
        {
            if ((index % context.Options.Source.CancellationCheckInterval) == 0)
                cancellationToken.ThrowIfCancellationRequested();
            signatures[index] = context.ComputeSignature(rows[index], cancellationToken);
            if (keys.Length != 0)
                keys[index] = context.ComputeKeySignature(rows[index], cancellationToken);
        }

        return new PreparedRows(rows, signatures, keys);
    }

    public static PreparedRows Select(
        PreparedRows source,
        IReadOnlyList<int> indices,
        int cancellationCheckInterval,
        CancellationToken cancellationToken)
    {
        var rows = new RowRecord[indices.Count];
        var signatures = new RowSignature[indices.Count];
        var keys = source.KeySignatures.Length == 0
            ? Array.Empty<RowKeySignature>()
            : new RowKeySignature[indices.Count];
        for (var index = 0; index < indices.Count; index++)
        {
            if ((index % cancellationCheckInterval) == 0)
                cancellationToken.ThrowIfCancellationRequested();
            var sourceIndex = indices[index];
            rows[index] = source.Rows[sourceIndex];
            signatures[index] = source.Signatures[sourceIndex];
            if (keys.Length != 0)
                keys[index] = source.KeySignatures[sourceIndex];
        }

        return new PreparedRows(rows, signatures, keys);
    }
}

internal readonly record struct AlignmentStep(int BaseOrdinal, int SourceOrdinal);

internal static class RowAligner
{
    private const byte Diagonal = 0;
    private const byte Delete = 1;
    private const byte Insert = 2;

    public static AlignmentStep[] Align(
        PreparedRows left,
        PreparedRows right,
        ComparisonContext context,
        CancellationToken cancellationToken)
    {
        var anchors = BuildAnchors(left, right, context, cancellationToken);
        var result = new List<AlignmentStep>(checked(left.Rows.Length + right.Rows.Length));
        var leftStart = 0;
        var rightStart = 0;
        foreach (var anchor in anchors)
        {
            AlignGap(
                left,
                leftStart,
                anchor.BaseOrdinal,
                right,
                rightStart,
                anchor.SourceOrdinal,
                context,
                result,
                cancellationToken);
            result.Add(anchor);
            leftStart = anchor.BaseOrdinal + 1;
            rightStart = anchor.SourceOrdinal + 1;
        }

        AlignGap(
            left,
            leftStart,
            left.Rows.Length,
            right,
            rightStart,
            right.Rows.Length,
            context,
            result,
            cancellationToken);
        return result.ToArray();
    }

    public static ChangeKind GetChangeKind(
        AlignmentStep step,
        PreparedRows left,
        PreparedRows right,
        ComparisonContext context,
        CancellationToken cancellationToken)
    {
        if (step.BaseOrdinal < 0)
            return ChangeKind.Added;
        if (step.SourceOrdinal < 0)
            return ChangeKind.Removed;
        return RowsExactlyEqual(left, step.BaseOrdinal, right, step.SourceOrdinal, context, cancellationToken)
            ? ChangeKind.Unchanged
            : ChangeKind.Modified;
    }

    public static bool RowsExactlyEqual(
        PreparedRows left,
        int leftIndex,
        PreparedRows right,
        int rightIndex,
        ComparisonContext context,
        CancellationToken cancellationToken)
    {
        return left.Signatures[leftIndex] == right.Signatures[rightIndex] &&
            context.RowsEqual(left.Rows[leftIndex], right.Rows[rightIndex], cancellationToken);
    }

    private static List<AlignmentStep> BuildAnchors(
        PreparedRows left,
        PreparedRows right,
        ComparisonContext context,
        CancellationToken cancellationToken)
    {
        List<AlignmentStep> keyAnchors;
        if (context.Options.KeyColumns.Length == 0)
        {
            keyAnchors = [];
        }
        else
        {
            keyAnchors = FindUniqueAnchors(
                left,
                0,
                left.Rows.Length,
                right,
                0,
                right.Rows.Length,
                useKeys: true,
                context,
                cancellationToken);
        }

        // Key anchors partition the sheet first; exact full-row anchors then stabilize each gap.
        var anchors = new List<AlignmentStep>();
        var leftStart = 0;
        var rightStart = 0;
        foreach (var keyAnchor in keyAnchors)
        {
            anchors.AddRange(FindUniqueAnchors(
                left,
                leftStart,
                keyAnchor.BaseOrdinal,
                right,
                rightStart,
                keyAnchor.SourceOrdinal,
                useKeys: false,
                context,
                cancellationToken));
            anchors.Add(keyAnchor);
            leftStart = keyAnchor.BaseOrdinal + 1;
            rightStart = keyAnchor.SourceOrdinal + 1;
        }

        anchors.AddRange(FindUniqueAnchors(
            left,
            leftStart,
            left.Rows.Length,
            right,
            rightStart,
            right.Rows.Length,
            useKeys: false,
            context,
            cancellationToken));
        return anchors;
    }

    private static List<AlignmentStep> FindUniqueAnchors(
        PreparedRows left,
        int leftStart,
        int leftEnd,
        PreparedRows right,
        int rightStart,
        int rightEnd,
        bool useKeys,
        ComparisonContext context,
        CancellationToken cancellationToken)
    {
        if (leftStart >= leftEnd || rightStart >= rightEnd)
            return [];

        var leftOccurrences = CountSignatures(
            left,
            leftStart,
            leftEnd,
            useKeys,
            context,
            cancellationToken);
        var rightOccurrences = CountSignatures(
            right,
            rightStart,
            rightEnd,
            useKeys,
            context,
            cancellationToken);
        var candidates = new List<AlignmentStep>();
        var inspected = 0;

        foreach (var pair in leftOccurrences)
        {
            if ((inspected++ % context.Options.Source.CancellationCheckInterval) == 0)
                cancellationToken.ThrowIfCancellationRequested();
            if (pair.Value.Count != 1 ||
                !rightOccurrences.TryGetValue(pair.Key, out var rightOccurrence) ||
                rightOccurrence.Count != 1)
            {
                continue;
            }

            var leftIndex = pair.Value.Index;
            var rightIndex = rightOccurrence.Index;
            var verified = useKeys
                ? context.KeysEqual(left.Rows[leftIndex], right.Rows[rightIndex], cancellationToken)
                : context.RowsEqual(left.Rows[leftIndex], right.Rows[rightIndex], cancellationToken);
            if (verified)
                candidates.Add(new AlignmentStep(leftIndex, rightIndex));
        }

        candidates.Sort(static (first, second) =>
        {
            var leftComparison = first.BaseOrdinal.CompareTo(second.BaseOrdinal);
            return leftComparison != 0
                ? leftComparison
                : first.SourceOrdinal.CompareTo(second.SourceOrdinal);
        });
        return LongestIncreasingSubsequence(candidates, context, cancellationToken);
    }

    private static Dictionary<RowSignature, Occurrence> CountSignatures(
        PreparedRows rows,
        int start,
        int end,
        bool useKeys,
        ComparisonContext context,
        CancellationToken cancellationToken)
    {
        var result = new Dictionary<RowSignature, Occurrence>();
        for (var index = start; index < end; index++)
        {
            if (((index - start) % context.Options.Source.CancellationCheckInterval) == 0)
                cancellationToken.ThrowIfCancellationRequested();
            RowSignature signature;
            if (useKeys)
            {
                var key = rows.KeySignatures[index];
                if (context.Options.Source.RequireNonBlankKey && !key.HasNonBlankValue)
                    continue;
                signature = key.Signature;
            }
            else
            {
                signature = rows.Signatures[index];
            }

            if (result.TryGetValue(signature, out var occurrence))
                result[signature] = new Occurrence(occurrence.Index, occurrence.Count + 1);
            else
                result.Add(signature, new Occurrence(index, 1));
        }

        return result;
    }

    private static List<AlignmentStep> LongestIncreasingSubsequence(
        List<AlignmentStep> candidates,
        ComparisonContext context,
        CancellationToken cancellationToken)
    {
        if (candidates.Count == 0)
            return [];

        var tails = new int[candidates.Count];
        var previous = new int[candidates.Count];
        Array.Fill(previous, -1);
        var length = 0;

        for (var candidateIndex = 0; candidateIndex < candidates.Count; candidateIndex++)
        {
            if ((candidateIndex % context.Options.Source.CancellationCheckInterval) == 0)
                cancellationToken.ThrowIfCancellationRequested();
            var sourceIndex = candidates[candidateIndex].SourceOrdinal;
            var low = 0;
            var high = length;
            while (low < high)
            {
                var middle = low + ((high - low) >> 1);
                if (candidates[tails[middle]].SourceOrdinal < sourceIndex)
                    low = middle + 1;
                else
                    high = middle;
            }

            if (low > 0)
                previous[candidateIndex] = tails[low - 1];
            tails[low] = candidateIndex;
            if (low == length)
                length++;
        }

        var result = new AlignmentStep[length];
        var current = tails[length - 1];
        for (var index = length - 1; index >= 0; index--)
        {
            result[index] = candidates[current];
            current = previous[current];
        }

        return [.. result];
    }

    private static void AlignGap(
        PreparedRows left,
        int leftStart,
        int leftEnd,
        PreparedRows right,
        int rightStart,
        int rightEnd,
        ComparisonContext context,
        List<AlignmentStep> result,
        CancellationToken cancellationToken)
    {
        while (leftStart < leftEnd && rightStart < rightEnd &&
            RowsExactlyEqual(left, leftStart, right, rightStart, context, cancellationToken))
        {
            result.Add(new AlignmentStep(leftStart++, rightStart++));
        }

        var suffixCount = 0;
        while (leftStart < leftEnd && rightStart < rightEnd &&
            RowsExactlyEqual(left, leftEnd - 1, right, rightEnd - 1, context, cancellationToken))
        {
            leftEnd--;
            rightEnd--;
            suffixCount++;
        }

        var leftCount = leftEnd - leftStart;
        var rightCount = rightEnd - rightStart;
        if (leftCount == 0)
        {
            for (var index = rightStart; index < rightEnd; index++)
                result.Add(new AlignmentStep(-1, index));
        }
        else if (rightCount == 0)
        {
            for (var index = leftStart; index < leftEnd; index++)
                result.Add(new AlignmentStep(index, -1));
        }
        else
        {
            var matrixCells = checked(((long)leftCount + 1) * (rightCount + 1));
            if (leftCount <= context.Options.Source.MaxAlignmentGapRows &&
                rightCount <= context.Options.Source.MaxAlignmentGapRows &&
                matrixCells <= context.Options.Source.MaxDynamicProgrammingCells &&
                matrixCells <= int.MaxValue)
            {
                AlignDynamic(
                    left,
                    leftStart,
                    leftEnd,
                    right,
                    rightStart,
                    rightEnd,
                    context,
                    result,
                    cancellationToken);
            }
            else
            {
                AlignWithLookahead(
                    left,
                    leftStart,
                    leftEnd,
                    right,
                    rightStart,
                    rightEnd,
                    context,
                    result,
                    cancellationToken);
            }
        }

        for (var offset = 0; offset < suffixCount; offset++)
            result.Add(new AlignmentStep(leftEnd + offset, rightEnd + offset));
    }

    private static void AlignDynamic(
        PreparedRows left,
        int leftStart,
        int leftEnd,
        PreparedRows right,
        int rightStart,
        int rightEnd,
        ComparisonContext context,
        List<AlignmentStep> result,
        CancellationToken cancellationToken)
    {
        // A byte-per-cell trace keeps Levenshtein alignment deterministic within the configured cap.
        var leftCount = leftEnd - leftStart;
        var rightCount = rightEnd - rightStart;
        var width = rightCount + 1;
        var trace = new byte[checked((leftCount + 1) * width)];
        var previous = new int[width];
        var current = new int[width];

        for (var column = 1; column <= rightCount; column++)
        {
            previous[column] = column;
            trace[column] = Insert;
        }

        long work = 0;
        for (var row = 1; row <= leftCount; row++)
        {
            current[0] = row;
            trace[row * width] = Delete;
            for (var column = 1; column <= rightCount; column++)
            {
                if ((work++ % context.Options.Source.CancellationCheckInterval) == 0)
                    cancellationToken.ThrowIfCancellationRequested();

                var equal = RowsExactlyEqual(
                    left,
                    leftStart + row - 1,
                    right,
                    rightStart + column - 1,
                    context,
                    cancellationToken);
                var diagonalCost = previous[column - 1] + (equal ? 0 : 1);
                var deleteCost = previous[column] + 1;
                var insertCost = current[column - 1] + 1;

                byte direction;
                int cost;
                if (diagonalCost <= deleteCost && diagonalCost <= insertCost)
                {
                    direction = Diagonal;
                    cost = diagonalCost;
                }
                else if (deleteCost <= insertCost)
                {
                    direction = Delete;
                    cost = deleteCost;
                }
                else
                {
                    direction = Insert;
                    cost = insertCost;
                }

                current[column] = cost;
                trace[row * width + column] = direction;
            }

            (previous, current) = (current, previous);
        }

        var reversed = new List<AlignmentStep>(leftCount + rightCount);
        var leftPosition = leftCount;
        var rightPosition = rightCount;
        while (leftPosition > 0 || rightPosition > 0)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var direction = trace[leftPosition * width + rightPosition];
            if (leftPosition > 0 && rightPosition > 0 && direction == Diagonal)
            {
                reversed.Add(new AlignmentStep(
                    leftStart + --leftPosition,
                    rightStart + --rightPosition));
            }
            else if (leftPosition > 0 && (rightPosition == 0 || direction == Delete))
            {
                reversed.Add(new AlignmentStep(leftStart + --leftPosition, -1));
            }
            else
            {
                reversed.Add(new AlignmentStep(-1, rightStart + --rightPosition));
            }
        }

        for (var index = reversed.Count - 1; index >= 0; index--)
            result.Add(reversed[index]);
    }

    private static void AlignWithLookahead(
        PreparedRows left,
        int leftStart,
        int leftEnd,
        PreparedRows right,
        int rightStart,
        int rightEnd,
        ComparisonContext context,
        List<AlignmentStep> result,
        CancellationToken cancellationToken)
    {
        // Oversized gaps resynchronize on nearby exact rows without quadratic storage.
        var leftIndex = leftStart;
        var rightIndex = rightStart;
        while (leftIndex < leftEnd && rightIndex < rightEnd)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (RowsExactlyEqual(left, leftIndex, right, rightIndex, context, cancellationToken))
            {
                result.Add(new AlignmentStep(leftIndex++, rightIndex++));
                continue;
            }

            var rightMatch = FindMatch(
                left,
                leftIndex,
                right,
                rightIndex + 1,
                rightEnd,
                context,
                cancellationToken);
            var leftMatch = FindMatch(
                right,
                rightIndex,
                left,
                leftIndex + 1,
                leftEnd,
                context,
                cancellationToken);
            var rightDistance = rightMatch < 0 ? int.MaxValue : rightMatch - rightIndex;
            var leftDistance = leftMatch < 0 ? int.MaxValue : leftMatch - leftIndex;

            if (rightDistance <= leftDistance && rightMatch >= 0)
            {
                while (rightIndex < rightMatch)
                    result.Add(new AlignmentStep(-1, rightIndex++));
            }
            else if (leftMatch >= 0)
            {
                while (leftIndex < leftMatch)
                    result.Add(new AlignmentStep(leftIndex++, -1));
            }
            else
            {
                result.Add(new AlignmentStep(leftIndex++, rightIndex++));
            }
        }

        while (leftIndex < leftEnd)
            result.Add(new AlignmentStep(leftIndex++, -1));
        while (rightIndex < rightEnd)
            result.Add(new AlignmentStep(-1, rightIndex++));
    }

    private static int FindMatch(
        PreparedRows needleRows,
        int needleIndex,
        PreparedRows searchRows,
        int searchStart,
        int searchEnd,
        ComparisonContext context,
        CancellationToken cancellationToken)
    {
        var boundedEnd = Math.Min(
            searchEnd,
            (int)Math.Min(
                int.MaxValue,
                (long)searchStart + context.Options.Source.AlignmentLookaheadRows));
        for (var index = searchStart; index < boundedEnd; index++)
        {
            if (RowsExactlyEqual(
                needleRows,
                needleIndex,
                searchRows,
                index,
                context,
                cancellationToken))
            {
                return index;
            }
        }

        return -1;
    }

    private readonly record struct Occurrence(int Index, int Count);
}
