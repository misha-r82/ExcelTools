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
        protected int _coluumn;
        public int Col { get { return _coluumn; } }
        public FilterSetter(FilterProto filter)
        {
            _filter = filter;            
        }
        
        public abstract void SetFilter(object criteria1, object criteria2);
        public abstract void RemoveFilter();
    }

    public class TableFilterSetter : FilterSetter
    {
        private Range _rng;
        public TableFilterSetter(FilterProto filter, bool isFilterCreated) : base(filter)
        {
            _rng = Current.CurRegion.ActiveCell;
            _coluumn = _rng.Column - Current.CurRegion.TblRange.FirstCol + 1;
            var cols = Current.CurRegion.ActiveRow.ExCells;
            if (isFilterCreated) return;
            filter.CanFilter = false;
            int i = 0;
            for (; i < cols.Length; i++)
            {
                if (!string.Equals(cols[i].ColName, filter.Name, StringComparison.OrdinalIgnoreCase)) continue;
                filter.CanFilter = true;
                _rng = cols[i].Rng;
                _coluumn = i + 1;
                break;
            }

        }
        public override void SetFilter(object criteria1, object criteria2)
        {

                try
                {
                    if (_filter.IsListMode)
                    {
                        if (_filter.SelectedValues == null || _filter.SelectedValues.Length == 0) return;
                        string[] strArr;
                        if (_filter is DateFilter)
                        {
                            var orig = Current.CurRegion.ActiveWs.Application.ActiveWindow.AutoFilterDateGrouping;
                            Current.CurRegion.ActiveWs.Application.ActiveWindow.AutoFilterDateGrouping = false;
                            strArr = _filter.SelectedValues.OfType<CellValue>()
                            .Select(v => v.ValDate.ToString(DateFilter.DATE_FORMAT)).ToArray();
                            _rng.CurrentRegion.AutoFilter(_coluumn, strArr, XlAutoFilterOperator.xlFilterValues,
                                Type.Missing, true);
                            Current.CurRegion.ActiveWs.Application.ActiveWindow.AutoFilterDateGrouping = orig;
                        }
                        else
                        {
                         strArr= _filter.SelectedValues.OfType<CellValue>().Select(v => v.ToString()).ToArray();
                        _rng.CurrentRegion.AutoFilter(_coluumn, strArr, XlAutoFilterOperator.xlFilterValues,
                                Type.Missing, true);
                        }

                        return;
                    }
                    if (criteria2 == null)
                        _rng.AutoFilter(_coluumn, criteria1);
                    else _rng.AutoFilter(_coluumn, criteria1, XlAutoFilterOperator.xlAnd, criteria2);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Не удалось установить фильтр!");
                }
        }
        public override void RemoveFilter()
        {
            try
            {
                _rng.AutoFilter(_coluumn);
            }
            catch (Exception e)
            {
                MessageBox.Show("Не удалось снять фильтр!");
            }
        }

    }
}
