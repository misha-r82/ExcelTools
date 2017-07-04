using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Office.Interop.Excel;

namespace ExcelTools.Filters
{
    public abstract class FilterSetter
    {
        protected FilterProto _filter;
        protected int _colnum;
        public int Col { get { return _colnum; } }
        public FilterSetter(FilterProto filter, int col)
        {
            _filter = filter;
            _colnum = col;
        }
        
        public abstract void SetFilter(object criteria1, object criteria2);
        public abstract void RemoveFilter();
    }

    public class TableFilterSetter : FilterSetter
    {
        private Range _rng;
        public TableFilterSetter(FilterProto filter, int col, Range rng) : base(filter, col)
        {
            _rng = rng;
        }
        // только при изменении диапазона фильтра
        public TableFilterSetter(FilterProto filter) : base(filter, filter.Setter.Col)
        {
            var cols = Current.CurRegion.ActiveRow.ExCells;
            filter.CanFilter = false;
            int i = 0;
            for (; i < cols.Length; i++)
            {
                if (!string.Equals(cols[i].ColName, filter.Name, StringComparison.OrdinalIgnoreCase)) continue;
                filter.CanFilter = true;
                _rng = cols[i].Rng;
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
                    if (_filter.IsListMode)
                    {
                        if (_filter.SelectedValues.Length == 0) return;
                        var strArr = _filter.SelectedValues.Select(v => v.ToString()).ToArray();
                        _rng.CurrentRegion.AutoFilter(_colnum, strArr, XlAutoFilterOperator.xlFilterValues,
                            Type.Missing, true);
                        return;
                    }
                    if (criteria2 == null)
                        _rng.AutoFilter(_colnum, criteria1);
                    else _rng.AutoFilter(_colnum, criteria1, XlAutoFilterOperator.xlAnd, criteria2);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Не удалось установить фильтр!");
                }

            }
        }
        public override void RemoveFilter()
        {
            try
            {
                _rng.AutoFilter(_colnum);
            }
            catch (Exception e)
            {
                MessageBox.Show("Не удалось снять фильтр!");
            }
        }

    }
    public class PivotFilterSetter : FilterSetter
    {
        private PivotField _pivField;
        public PivotFilterSetter(FilterProto filter) : base(filter, -1)
        {
            filter.CanFilter = false;
            int i = 0;
            var fields = Current.CurRegion.ActiveRow.PivotFields;
            for (; i < fields.Length; i++)
            {
                if (!string.Equals(fields[i].Name, filter.Name, StringComparison.OrdinalIgnoreCase)) continue;
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
    //            ActiveSheet.PivotTables("Сводная таблица1").PivotFields("Клиент").PivotFilters._
    //    Add2 Type:= xlCaptionContains, Value1:= "ва"
    //ActiveSheet.PivotTables("Сводная таблица1").PivotFields("Клиент")._
    //    ClearAllFilters
    //ActiveSheet.PivotTables("Сводная таблица1").PivotFields("Дата покупки")._
    //    PivotFilters.Add2 Type:= xlDateBetween, Value1:= "01.12.2016", Value2:= _
    //    "08.12.2016"
                /*try
                {
                    if (_filter.IsListMode)
                    {
                        if (_filter.SelectedValues.Length == 0) return;
                        var strArr = _filter.SelectedValues.Select(v => v.ToString()).ToArray();
                        _rng.CurrentRegion.AutoFilter(_colnum, strArr, XlAutoFilterOperator.xlFilterValues,
                            Type.Missing, true);
                        return;
                    }
                    if (criteria2 == null)
                        _rng.AutoFilter(_colnum, criteria1);
                    else _rng.AutoFilter(_colnum, criteria1, XlAutoFilterOperator.xlAnd, criteria2);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Не удалось установить фильтр!");
                }*/
                _pivField.PivotFilters.Add(XlPivotFilterType.xlCaptionContains, Type.Missing, criteria1);

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
