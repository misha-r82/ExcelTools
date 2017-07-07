using System;
using System.Diagnostics;
using System.Linq;
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

        private void SetListFilter()
        {
            var tmp = (Excel.PivotItems)_pivField.PivotItems();
            /*foreach (dynamic item in _pivField.PivotItems())
            {
                var pivItm = item as PivotItem;
                pivItm.Visible = _filter.ValueList.Any(v => v.StrVal == pivItm.Value);
            }  */
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
                    Debug.WriteLine("�� ������� ���������� PivotFilter " + e.Message);
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
                MessageBox.Show("�� ������� ����� ������!");
            }
        }

    }
}