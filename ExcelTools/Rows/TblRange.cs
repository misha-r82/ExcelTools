using Microsoft.Office.Interop.Excel;

namespace ExcelTools
{
    public class TblRange
    {
        private int _firstRow;
        private int _lastRow;
        private int _firstCol;
        private int _lastCol;
        private Range _curRng;

        public int FirstRow => _firstRow;

        public int LastRow => _lastRow;

        public int FirstCol => _firstCol;

        public int LastCol => _lastCol;

        public Range CurRng => _curRng;

        public TblRange(Range cell)
        {
            _curRng = cell.CurrentRegion;
            Range firstCell = (Range)_curRng[2, 1];
            _firstRow = firstCell.Row;
            _firstCol = firstCell.Column;
            _lastRow = _firstRow + _curRng.Rows.Count - 2;
            _lastCol = _firstCol + _curRng.Columns.Count;
        }
    }
}