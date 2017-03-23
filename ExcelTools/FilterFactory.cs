using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Office.Interop.Excel;

namespace ExcelTools
{
    static class FilterFactory
    {
 
        public static FilterProto CreateFilter()
        {
            Range rng = Current.CurRegion.ActiveCell;
            var activeCell = new Cell(rng, true);
            
            switch (activeCell.Type)
            {
                case CellTypes.str: return new StrFilter(activeCell);
                case CellTypes.numeric: return new NumericFilter(activeCell);
                case CellTypes.date: return new DateFilter(activeCell);
                case CellTypes.time: return new TimeFilter(activeCell);
                default: return new StrFilter(activeCell);

            }
            //
            //if (rng.v)


        }
    }
}
