using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Windows;
using Microsoft.Office.Interop.Excel;

namespace ExcelTools.Filters
{
    public class PivotFilterSetter : FilterSetter
    {
        public static string TIME_FORMAT = @"hh:mm";
        public static string DATE_FORMAT { get; set; } = "dd.MM.yyyy";
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

        private void SetListFilter()
        {
            var items = (PivotItems)_pivField.PivotItems();
            var selValues = new string[0];
            //var pivRng = (Range)_pivField.DataRange;
            ////var fltVals = _filter.SelectedValues.OfType<CellValue>().Select(v => v.XlVal).ToArray();
            //foreach (object cellobj in pivRng.Cells)
            //{
            //    Range cellRng = (Range) cellobj;
            //    var cell = new CellValue(cellRng);
            //    Debug.WriteLine(cellRng.NumberFormat + " - " + cellRng.Value2);
            //    //if (fltVals.Contains(cell.XlVal))
            //    //    selValues.Add(cellRng.ToString());
            //}
            if (_filter.GetType() == typeof(StrFilter) || _filter.GetType() == typeof(NumericFilter))
            {
                selValues = _filter.SelectedValues.Select(v => v.ToString()).ToArray();
            }
            else if (_filter.GetType() == typeof(DateFilter))
            {
                selValues = _filter.SelectedValues.OfType<CellValue>()
                    .Select(v=>v.ValDate.ToString(DATE_FORMAT, CultureInfo.InvariantCulture)).ToArray();
            }
            else if (_filter.GetType() == typeof(TimeFilter))
            {
                selValues = _filter.SelectedValues.OfType<CellValue>()
                    .Select(v=>v.ValTime.ToString(TIME_FORMAT, CultureInfo.InvariantCulture)).ToArray();
            }
            
            foreach (dynamic item in items)
            {
                var pivItm = item as PivotItem;
                pivItm.Visible = selValues.Contains(pivItm.Value);
            }  
        }
        public override void SetFilter(object criteria1, object criteria2)
        {
            RemoveFilter();
             try
                {
                    if (_filter.IsListMode)
                    {
                        SetListFilter();
                        return; 
                    }
                    if (_filter.GetType() == typeof(StrFilter))

                            _pivField.PivotFilters.Add(XlPivotFilterType.xlCaptionContains, Type.Missing, criteria1);
                        else SetListFilter();
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