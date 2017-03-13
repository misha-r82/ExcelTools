using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace ExcelTools
{
    public enum CellTypes
    {
        str,
        numeric, 
        date,
        time

    }

    public class ActiveRow : IEnumerable<Cell>
    {
        public Cell[] Cells;
        public Range RowRng { get; set; }
        public ActiveRow(Range rng)
        {
            RowRng = ExTools.RowByCell(rng);
            Cells = new Cell[RowRng.Count];
            int i = 0;
            foreach (object cell in RowRng.Cells)
                Cells[i++] = new Cell((Range) cell, true);
        }

        public IEnumerator<Cell> GetEnumerator()
        {
            return (IEnumerator<Cell>) Cells.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return Cells.GetEnumerator();
        }
    }
}
