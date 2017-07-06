using System;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows;
using ExcelTools.Annotations;
using ExcelTools.Filters;
using Microsoft.Office.Interop.Excel;

namespace ExcelTools
{
    public abstract class FilterProto : INotifyPropertyChanged, IEquatable<FilterProto>
    {
        public string Name { get; }
        public CellValue[] ValueList { get; set; }
        private object[] _selectedValues;
  protected virtual object Criteria1 { get; }
        protected virtual object Criteria2 { get; }
        public FilterSetter Setter { get; set; }
        private bool _enabled;
        private bool _isListMode;
        private bool _canFilter;

        public abstract string Caption { get; }
        public object[] SelectedValues
        {
            get { return _selectedValues; }
            set
            {
                _selectedValues = value;
                OnPropertyChanged();
            }
        }
        public bool IsListMode
        {
            get { return _isListMode; }
            set
            {
                if (value == _isListMode) return;
                _isListMode = value;
                OnPropertyChanged();
            }
        }

        public bool Enabled
        {
            get { return _enabled; }
            set
            {
                if (_enabled == value) return;
                _enabled = value;
                OnPropertyChanged();
                if (!_enabled) RemoveFilter();
            }
        }
        public FilterProto()
        {
            var FilterRng = Current.CurRegion.ActiveCell;
            int col = FilterRng.Column - Current.CurRegion.firstCol;
            Setter = new TableFilterSetter(this, true);
            Name = Current.CurRegion.ActiveRow.ExCells[col].ColName;
            var tmpCell = new ExCell(FilterRng, true);
            ValueList = tmpCell.ValList;
            _canFilter = true;
            _enabled = true;
        }
        public void SetFilter()
        {
            if (CanFilter && Enabled)
            {
                RemoveFilter();
                Setter.SetFilter(Criteria1, Criteria2);
            }
        }
        public void RemoveFilter()

        {
            if (CanFilter)
                Setter.RemoveFilter();
        }      
        public void OnRangeChange()
        {
            if (Current.CurRegion.ActiveRow.PivotFields != null &&
                Current.CurRegion.ActiveRow.PivotFields.Length > 0)
                Setter = new PivotFilterSetter(this);
            else Setter = new TableFilterSetter(this, false);
            if (!CanFilter) return;
            SetFilter();
        }

        public bool CanFilter
        {
            get { return _canFilter; }
            set
            {
                if (value == _canFilter) return;
                _canFilter = value;
                OnPropertyChanged();
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        [NotifyPropertyChangedInvocator]
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            SetFilter();
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public bool Equals(FilterProto other)
        {
            if (ReferenceEquals(null, other)) return false;
            if (ReferenceEquals(this, other)) return true;
            return string.Equals(Name, other.Name) && Equals(Criteria1, other.Criteria1) && Equals(Criteria2, other.Criteria2);
        }

        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj)) return false;
            if (ReferenceEquals(this, obj)) return true;
            if (obj.GetType() != this.GetType()) return false;
            return Equals((FilterProto) obj);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                var hashCode = Name != null ? Name.GetHashCode() : 0;
                hashCode = (hashCode * 397) ^ (Criteria1 != null ? Criteria1.GetHashCode() : 0);
                hashCode = (hashCode * 397) ^ (Criteria2 != null ? Criteria2.GetHashCode() : 0);
                return hashCode;
            }
        }
    }
}