using ExcelMerge.Domain;

namespace ExcelMerge.Storage;

internal sealed class RowCache
{
    private readonly int _capacity;
    private readonly long _byteLimit;
    private readonly Dictionary<long, CacheEntry> _entries = new();
    private readonly LinkedList<long> _lru = new();
    private readonly object _sync = new();
    private long _sizeBytes;

    public RowCache(int capacity, long byteLimit)
    {
        _capacity = capacity;
        _byteLimit = byteLimit;
    }

    public int Count
    {
        get
        {
            lock (_sync)
            {
                return _entries.Count;
            }
        }
    }

    public long SizeBytes
    {
        get
        {
            lock (_sync)
            {
                return _sizeBytes;
            }
        }
    }

    public bool TryGet(long rowOrdinal, out RowRecord row)
    {
        lock (_sync)
        {
            if (!_entries.TryGetValue(rowOrdinal, out var entry))
            {
                row = default;
                return false;
            }

            _lru.Remove(entry.Node);
            _lru.AddFirst(entry.Node);
            row = entry.Row;
            return true;
        }
    }

    public void Add(long rowOrdinal, RowRecord row)
    {
        if (_capacity == 0 || _byteLimit == 0)
        {
            return;
        }

        var sizeBytes = EstimateSize(row);

        lock (_sync)
        {
            if (_entries.TryGetValue(rowOrdinal, out var existing))
            {
                Remove(existing);
            }

            if (sizeBytes > _byteLimit)
            {
                return;
            }

            var node = _lru.AddFirst(rowOrdinal);
            _entries.Add(rowOrdinal, new CacheEntry(row, node, sizeBytes));
            _sizeBytes += sizeBytes;

            while (_entries.Count > _capacity || _sizeBytes > _byteLimit)
            {
                var leastRecent = _lru.Last;
                if (leastRecent is null)
                {
                    break;
                }

                Remove(_entries[leastRecent.Value]);
            }
        }
    }

    public void Clear()
    {
        lock (_sync)
        {
            _entries.Clear();
            _lru.Clear();
            _sizeBytes = 0;
        }
    }

    private void Remove(CacheEntry entry)
    {
        _entries.Remove(entry.Node.Value);
        _lru.Remove(entry.Node);
        _sizeBytes -= entry.SizeBytes;
    }

    private static long EstimateSize(RowRecord row)
    {
        const long RowOverhead = 64;
        const long CellOverhead = 80;
        const long StringOverhead = 24;

        long size = RowOverhead + (CellOverhead * row.CellCount);
        var cells = row.Cells.Span;
        for (var i = 0; i < cells.Length; i++)
        {
            var value = cells[i].Value;
            size += EstimateString(value.DisplayText);
            size += EstimateString(value.Formula);

            var scalar = value.IsFormula ? value.CachedValue : value.Scalar;
            if (scalar is { Kind: CellKind.Text or CellKind.Error } textScalar)
            {
                size += EstimateString(textScalar.TextValue);
            }
        }

        return size;

        static long EstimateString(string? value) =>
            value is null ? 0 : StringOverhead + ((long)value.Length * sizeof(char));
    }

    private sealed record CacheEntry(
        RowRecord Row,
        LinkedListNode<long> Node,
        long SizeBytes);
}
