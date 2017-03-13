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
            Range curRng = cell.CurrentRegion;
            int colCount = curRng.Columns.Count;
            int row = cell.Row;
            Debug.WriteLine(((Range)curRng.Cells[row, 1]).Address + " " + ((Range)curRng.Cells[row, 1]).Address);
            var ws = Current.CurRegion.ActiveWs;
            var rowRng = ws.Range[ws.Cells[row, 1], ws.Cells[row, colCount]];
            return rowRng;
        }
    }
}
