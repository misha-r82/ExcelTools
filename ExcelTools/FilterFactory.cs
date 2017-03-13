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
 
        public static FilterProto CreateFilter(Range rng)
        {
            if (rng.Columns.Count > 1)
            {
                MessageBox.Show("Допускается выбор только одного сталбца");
                return null;
            }
            object val = rng.Count == 1 ? rng.Value : ((Range)rng.Cells[1,0]).Value;
            if (val is string) return new StrFilter(rng);
            if (val is double) return new NumericFilter(rng);
            if (val is DateTime) return new DateTimeFilter(rng);
            return null;
            //
            //if (rng.v)


        }
    }
}
