using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Office.Interop.Excel;
using System.Collections.ObjectModel;

namespace ExcelTools
{
    public static class FilterCollection
    {
        static FilterCollection()
        {
            Filters = new ObservableCollection<FilterProto>();
        }
        public static ObservableCollection<FilterProto> Filters { get; }

        private static FilterProto CreateFilter(Cell activeCell)
        {           
            switch (activeCell.Type)
            {
                case CellTypes.str: return new StrFilter(activeCell);
                case CellTypes.numeric: return new NumericFilter(activeCell);
                case CellTypes.date: return new DateFilter(activeCell);
                case CellTypes.time: return new TimeFilter(activeCell);
                default: return new StrFilter(activeCell);
            }
        }

        public static void Remove(FilterProto filter)
        {
            filter.RemoveFilter();
            Filters.Remove(filter);
        }
        
        public static void AddFilter(Range activeCell)
        {
            var cell = new Cell(activeCell, true);
            var flt = CreateFilter(cell);
            if (flt != null && !Filters.Contains(flt)) Filters.Add(flt);
        }
    }

}
