using System;
using System.Text.RegularExpressions;

namespace ExcelTools
{
    public sealed class CellReference : IComparable<CellReference>
    {
        private const int AlphabetCount = 26;

        private static readonly Regex s_columnNameRegex = new Regex("[A-Z]+");
        private static readonly Regex s_rowIndexReges = new Regex(@"\d+");

        public string ColumnName { get; }

        public int ColumnIndex => _lazyColumnIndex.Value;

        public int RowIndex { get; }

        private readonly Lazy<int> _lazyColumnIndex;

        public static CellReference Parse(string str)
        {
            var columnNameMatch = s_columnNameRegex.Match(str);
            var rowIndexMatch = s_rowIndexReges.Match(str);

            if (!columnNameMatch.Success || !rowIndexMatch.Success)
            {
                throw new ArgumentException("Expected reference of form \"A1\"");
            }

            return new CellReference(columnNameMatch.Value, int.Parse(rowIndexMatch.Value));
        }

        public static CellReference Create(int columnIndex, int rowIndex)
        {
            if (columnIndex <= 0)
            {
                throw new ArgumentOutOfRangeException(nameof(columnIndex), "Column index should be greater than zero");
            }

            return new CellReference(GetColumnName(columnIndex), rowIndex);
        }

        private CellReference(string columnName, int rowIndex)
        {
            if (rowIndex <= 0)
            {
                throw new ArgumentOutOfRangeException(nameof(rowIndex), "Row index should be greater than zero");
            }

            ColumnName = columnName;
            RowIndex = rowIndex;
            _lazyColumnIndex = new Lazy<int>(() => GetColumnIndex(columnName));
        }

        private static int GetColumnIndex(string columnName)
        {
            int index = 0;
            for (var i = 0; i < columnName.Length; ++i)
            {
                index = (index * AlphabetCount) + (columnName[i] - ('A' - 1));
            }
            return index;
        }

        private static string GetColumnName(int columnIndex)
        {
            string name = "";

            int div = columnIndex;
            int mod = 0;
            while (div > 0)
            {
                mod = (div - 1) % AlphabetCount;
                name = (char)('A' + mod) + name;
                div = (div - mod) / AlphabetCount;
            }

            return name;
        }

        public override string ToString()
        {
            return $"{ColumnName}{RowIndex}";
        }

        public int CompareTo(CellReference other)
        {
            if (other == null)
            {
                return 1;
            }

            if (RowIndex < other.RowIndex)
            {
                return -1;
            }
            else if (RowIndex > other.RowIndex)
            {
                return 1;
            }
            else if (ColumnIndex < other.ColumnIndex)
            {
                return -1;
            }
            else if (ColumnIndex > other.ColumnIndex)
            {
                return 1;
            }
            else
            {
                return 0;
            }
        }

        public static bool operator >(CellReference c1, CellReference c2)
        {
            return c1.CompareTo(c2) == 1;
        }

        public static bool operator <(CellReference c1, CellReference c2)
        {
            return c1.CompareTo(c2) == -1;
        }

        public static bool operator >=(CellReference c1, CellReference c2)
        {
            return c1.CompareTo(c2) >= 0;
        }

        public static bool operator <=(CellReference c1, CellReference c2)
        {
            return c1.CompareTo(c2) <= 0;
        }
    }
}
