using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Office.Interop.Excel;
using System.Collections.ObjectModel;
using System.Collections.Specialized;

namespace ExcelTools
{
    public static class FilterCollection
    {
        static FilterCollection()
        {
            Filters = new ObservableCollection<FilterProto>();
        }



        public static ObservableCollection<FilterProto> Filters { get; }

        private static FilterProto CreateFilter(ExCell activeExCell)
        {
            switch (activeExCell.Value.Type)
            {
                case CellValue.CellValType.String:
                    return new StrFilter(activeExCell);
                case CellValue.CellValType.Numeric:
                    return new NumericFilter(activeExCell);
                case CellValue.CellValType.Date:
                    return new DateFilter(activeExCell);
                case CellValue.CellValType.Time:
                    return new TimeFilter(activeExCell);
                default:
                    return new StrFilter(activeExCell);
            }
        }

        private static void Refilter()
        {
            foreach (var filter in Filters)
                filter.SetFilter();
        }

        public static void Remove(FilterProto filter)
        {
            filter.RemoveFilter();
            Filters.Remove(filter);
        }
        
        public static void AddFilter(Range activeCell)
        {
            var cell = new ExCell(activeCell, true);
            var flt = CreateFilter(cell);
            if (flt != null && !Filters.Contains(flt)) Filters.Add(flt);
        }
    }

}
