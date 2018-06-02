using System;

namespace ExcelTools
{
    public sealed class TableReference
    {
        public CellReference StartCell { get; }

        public CellReference EndCell { get; }

        public static TableReference Parse(string reference)
        {
            var strs = reference.Split(':');
            if (strs.Length != 2)
            {
                throw new ArgumentException("Expected reference of form \"A1:B2\"");
            }

            return new TableReference(CellReference.Parse(strs[0]), CellReference.Parse(strs[1]));
        }

        public TableReference(CellReference startCell, CellReference endCell)
        {
            if (startCell == null)
            {
                throw new ArgumentNullException(nameof(startCell));
            }
            if (endCell == null)
            {
                throw new ArgumentNullException(nameof(endCell));
            }

            StartCell = startCell;
            EndCell = endCell;
        }

        public override string ToString()
        {
            return $"{StartCell}:{EndCell}";
        }
    }
}
