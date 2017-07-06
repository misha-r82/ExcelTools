using System;
using System.Diagnostics;
using System.Windows;
using Microsoft.Office.Interop.Excel;

namespace ExcelTools.Filters
{
    public class PivotFilterSetter : FilterSetter
    {
        private PivotField _pivField;
        public PivotFilterSetter(FilterProto filter) : base(filter)
        {
            filter.CanFilter = false;
            int i = 0;
            var fields = Current.CurRegion.ActiveRow.PivotFields;
            for (; i < fields.Length; i++)
            {
                if (!string.Equals(fields[i].Name, filter.Name, StringComparison.OrdinalIgnoreCase)) continue;
                /*if (filter.GetType() == typeof(DateFilter) || filter.GetType() == typeof(DateFilter))
                    if (fields[i].DataType != XlPivotFieldDataType.xlDate) continue;*/
                if (filter.GetType() == typeof(NumericFilter))
                    if (fields[i].DataType != XlPivotFieldDataType.xlNumber) continue;
                filter.CanFilter = true;
                _pivField = fields[i];
                break;
            }
        }
        public override void SetFilter(object criteria1, object criteria2)
        {
            if (_filter.CanFilter && _filter.Enabled)
            {
                RemoveFilter();
                try
                {
                    if (_filter.GetType() == typeof(StrFilter))
                        _pivField.PivotFilters.Add(XlPivotFilterType.xlCaptionContains, Type.Missing, criteria1);
                    if (_filter.GetType() == typeof(DateFilter))
                    {
                        var flt = (DateFilter) _filter;
                        var cultureinfo = System.Globalization.CultureInfo.InvariantCulture;
                        _pivField.PivotFilters.Add(XlPivotFilterType.xlDateBetween, Type.Missing,
                            flt.From.ToString(CellValue.DATE_FORMAT, cultureinfo),
                            flt.To.ToString(CellValue.DATE_FORMAT, cultureinfo));
                    }
                    if (_filter.GetType() == typeof(TimeFilter))
                    {
                        var flt = (TimeFilter) _filter;
                        var cultureinfo = System.Globalization.CultureInfo.InvariantCulture;
                        _pivField.PivotFilters.Add(XlPivotFilterType.xlDateBetween, Type.Missing,
                            flt.From.ToString(CellValue.TIME_FORMAT, cultureinfo),
                            flt.To.ToString(CellValue.TIME_FORMAT, cultureinfo));
                    }
                    if (_filter.GetType() == typeof(NumericFilter))
                    {
                        var flt = (NumericFilter) _filter;
                        _pivField.PivotFilters.Add(XlPivotFilterType.xlCaptionIsBetween, Type.Missing,
                            flt.From.ToString(), flt.To.ToString());
                    }
                }
                catch (Exception e)
                {
                    Debug.WriteLine("не удалось установить PivotFilter " + e.Message);
                }
            }
        }
        public override void RemoveFilter()
        {
            try
            {
                _pivField.ClearAllFilters();
                //_rng.AutoFilter(_colnum);
            }
            catch (Exception e)
            {
                MessageBox.Show("Не удалось снять фильтр!");
            }
        }

    }
}