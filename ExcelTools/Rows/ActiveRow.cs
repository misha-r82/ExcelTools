using System;
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
        public PivotField[] PivotFields;
        public Range RowRng { get; set; }
        public ActiveRow(Range rng)
        {
            int i = 0;
            if (Current.CurRegion.IsTableCell)
            {
                PivotFields = new PivotField[0];
                RowRng = ExTools.RowByCell(rng);
                ExCells = new ExCell[RowRng.Count];
                foreach (object cell in RowRng.Cells)
                    ExCells[i++] = new ExCell((Range) cell, true);                
            }
            else
            {
                
                try
                {
                    ExCells = new ExCell[0];
                    var fields = (PivotFields)Current.CurRegion.PivotTable.PivotFields();
                    PivotFields = new PivotField[fields.Count];
                    for (; i <= fields.Count; i++)
                        PivotFields[i] = (PivotField) fields.Item(i);
                }
                catch (Exception e) 
                { }                
            }

            

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
