using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace ExcelTools
{
    static class ExTools
    {
        public static Range RowByCell(Range cell)
        {
            if (cell.Count != 1) throw new ArgumentException("Попытка создать ActiveRow из диапазона, содержащего не 1 ячейку!");
            if (!Current.CurRegion.IsTableCell)
            {
                
            }
            Range curRng = cell.CurrentRegion;
            int colCount = curRng.Columns.Count;
            int row = cell.Row;
            var ws = Current.CurRegion.ActiveWs;
            int firstCol = curRng.Column;
            Range rowRng = ws.Range[ws.Cells[row, firstCol], ws.Cells[row, firstCol + colCount - 1]];
            return rowRng;
        }
    }
}
