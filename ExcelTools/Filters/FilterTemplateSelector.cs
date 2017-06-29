using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;

namespace ExcelTools
{
    public class FilterTemplateSelector : DataTemplateSelector
    {

            public override DataTemplate SelectTemplate(object item, DependencyObject container)
            {
                ContentPresenter pres = container as ContentPresenter;
                if(item is StrFilter) return pres.FindResource("StrFilterTemplate") as DataTemplate;
                if (item is NumericFilter) return pres.FindResource("NumericFilterTemplate") as DataTemplate;
                if (item is DateFilter) return pres.FindResource("DateFilterTemplate") as DataTemplate;
                if (item is TimeFilter) return pres.FindResource("TimeFilterTemplate") as DataTemplate;
                return null;
            }
        
    }
    public class ActiveRowTemplateSelector : DataTemplateSelector
    {

        public override DataTemplate SelectTemplate(object item, DependencyObject container)
        {
            ContentPresenter pres = container as ContentPresenter;
            var cell = item as ExCell;
            if (cell == null) return null;
            switch (cell.Value.Type) 
            {
                case CellValue.CellValType.String: return pres.FindResource("StrCellTemplate") as DataTemplate;
                case CellValue.CellValType.Numeric: return pres.FindResource("StrCellTemplate") as DataTemplate; 
                case CellValue.CellValType.Date: return pres.FindResource("DateTimeCellTemplate") as DataTemplate;
                case CellValue.CellValType.Time: return pres.FindResource("TimeCellTemplate") as DataTemplate;
            }
            return null;
        }

    }
}
