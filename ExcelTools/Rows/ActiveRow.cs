using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace ExcelTools
{
    public class ActiveRow : IEnumerable<ExCell>
    {
        public ExCell[] ExCells;
        public Range RowRng { get; set; }
        public ActiveRow(Range rng)
        {
            RowRng = ExTools.RowByCell(rng);
            ExCells = new ExCell[RowRng.Count];
            int i = 0;
            foreach (object cell in RowRng.Cells)
                ExCells[i++] = new ExCell((Range) cell, true);
        }

        public IEnumerator<ExCell> GetEnumerator()
        {
            return (IEnumerator<ExCell>) ExCells.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return ExCells.GetEnumerator();
        }
    }
}
