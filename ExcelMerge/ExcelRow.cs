using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelMerge
{
    public class ExcelRow : IEquatable<ExcelRow>
    {
        public int Index { get; private set; }
        public List<ExcelCell> Cells { get; private set; }

        public ExcelRow(int index, IEnumerable<ExcelCell> cells)
        {
            Index = index;
            Cells = cells.ToList();
        }

        private ExcelRow(int index, List<ExcelCell> cells)
        {
            Index = index;
            Cells = cells;
        }

        internal static ExcelRow FromCells(int index, List<ExcelCell> cells)
        {
            return new ExcelRow(index, cells);
        }

        public override bool Equals(object obj)
        {
            var other = obj as ExcelRow;

            return Equals(other);
        }

        public override int GetHashCode()
        {
            var hash = 7;
            foreach (var cell in Cells)
            {
                hash = hash * 13 + cell.Value.GetHashCode();
            }

            return hash;
        }

        public bool Equals(ExcelRow other)
        {
            if (other == null)
                return false;

            return GetHashCode() == other.GetHashCode();
        }

        public bool IsBlank()
        {
            return Cells.All(c => string.IsNullOrEmpty(c.Value));
        }

        public void UpdateCells(IEnumerable<ExcelCell> cells)
        {
            Cells = cells.ToList();
        }
    }

    internal class RowComparer : IEqualityComparer<ExcelRow>
    {
        public HashSet<int> IgnoreColumns { get; private set; }
        private readonly Dictionary<ExcelRow, int> hashCodes =
            new Dictionary<ExcelRow, int>(ReferenceEqualityComparer.Instance);

        public RowComparer(HashSet<int> ignoreColumns)
        {
            IgnoreColumns = ignoreColumns;
        }

        public bool Equals(ExcelRow x, ExcelRow y)
        {
            return GetHashCode(x).Equals(GetHashCode(y));
        }

        public int GetHashCode(ExcelRow obj)
        {
            int cached;
            if (hashCodes.TryGetValue(obj, out cached))
                return cached;

            var hash = 7;
            var index = 0;
            foreach (var cell in obj.Cells)
            {
                if (IgnoreColumns.Contains(index))
                    continue;

                hash = hash * 13 + cell.Value.GetHashCode();

                index++;
            }

            hashCodes.Add(obj, hash);
            return hash;
        }
    }
}
