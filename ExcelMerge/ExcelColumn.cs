using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelMerge
{
    public class ExcelColumn : IEquatable<ExcelColumn>
    {
        public List<ExcelCell> Cells { get; private set; }
        public int HeaderIndex { get; set; }

        public ExcelColumn()
        {
            Cells = new List<ExcelCell>();
        }

        public ExcelColumn(IEnumerable<ExcelCell> cells)
        {
            Cells = cells.ToList();
        }

        public override bool Equals(object obj)
        {
            var other = obj as ExcelColumn;

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

        public bool Equals(ExcelColumn other)
        {
            if (other == null)
                return false;

            return GetHashCode() == other.GetHashCode();
        }

        public bool IsBlank()
        {
            return Cells.All(c => string.IsNullOrEmpty(c.Value));
        }
    }

    internal class HeaderComparer : IEqualityComparer<ExcelColumn>
    {
        public bool Equals(ExcelColumn x, ExcelColumn y)
        {
            var valueX = x.HeaderIndex >= 0 && x.HeaderIndex < x.Cells.Count
                ? x.Cells[x.HeaderIndex].Value
                : string.Empty;
            var valueY = y.HeaderIndex >= 0 && y.HeaderIndex < y.Cells.Count
                ? y.Cells[y.HeaderIndex].Value
                : string.Empty;

            return valueX.Equals(valueY);
        }

        public int GetHashCode(ExcelColumn obj)
        {
            return obj.HeaderIndex >= 0 && obj.HeaderIndex < obj.Cells.Count
                ? obj.Cells[obj.HeaderIndex].Value.GetHashCode()
                : string.Empty.GetHashCode();
        }
    }
}
