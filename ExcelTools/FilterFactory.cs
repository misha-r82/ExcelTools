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
            var activeCell = new Cell(rng, false);
            
            switch (activeCell.Type)
            {
                case CellTypes.str: return new StrFilter(rng);
                case CellTypes.numeric: return new NumericFilter(rng);
                case CellTypes.date: return new DateFilter(rng);
                case CellTypes.time: return new TimeFilter(rng);
                default: return new StrFilter(rng);

            }
            //
            //if (rng.v)


        }
    }
}
