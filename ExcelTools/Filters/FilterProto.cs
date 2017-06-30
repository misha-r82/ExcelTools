using System;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows;
using ExcelTools.Annotations;
using Microsoft.Office.Interop.Excel;

namespace ExcelTools
{
    public abstract class FilterProto : INotifyPropertyChanged, IEquatable<FilterProto>
    {
        private int ColNum { get; }
        public virtual string Name { get; }
        public object[] ValueList { get; set; }
        private object[] _selectedValues;

        public object[] SelectedValues
        {
            get { return _selectedValues; }
            set
            {
                _selectedValues = value;
                OnPropertyChanged();
            }
        }

        private Range FilterRng { get; set; }
        protected virtual object Criteria1 { get; }
        protected virtual object Criteria2 { get; }
        private bool _enabled;
        private bool _isListMode;

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
            FilterRng = Current.CurRegion.ActiveCell;
            int col = FilterRng.Column - Current.CurRegion.firstCol;
            Name = Current.CurRegion.ActiveRow.ExCells[col].ColName;
            var tmpCell = new ExCell(FilterRng, true);
            ValueList = tmpCell.ValList;
            ColNum =  col + 1;
            _enabled = true;
        }
        public void SetFilter()
        {
            if (Enabled)
            {
                RemoveFilter();
                try
                {
                    if (IsListMode)
                    {
                        if (SelectedValues.Length == 0) return;
                        var strArr = SelectedValues.Select(v => v.ToString()).ToArray();
                        FilterRng.CurrentRegion.AutoFilter(ColNum, strArr, XlAutoFilterOperator.xlFilterValues,
                            Type.Missing, true);
                        return;
                    }
                    if (Criteria2 == null) FilterRng.AutoFilter(ColNum, Criteria1);   
                    else FilterRng.AutoFilter(ColNum, Criteria1, XlAutoFilterOperator.xlAnd, Criteria2);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Не удалось установить фильтр!");
                }

            }
        }
        public void RemoveFilter()
        {
            try
            {
                FilterRng.AutoFilter(ColNum);
            }
            catch (Exception e)
            {
                MessageBox.Show("Не удалось снять фильтр!");
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
            return ColNum == other.ColNum && string.Equals(Name, other.Name) && Equals(Criteria1, other.Criteria1) && Equals(Criteria2, other.Criteria2);
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
                var hashCode = ColNum;
                hashCode = (hashCode * 397) ^ (Name != null ? Name.GetHashCode() : 0);
                hashCode = (hashCode * 397) ^ (Criteria1 != null ? Criteria1.GetHashCode() : 0);
                hashCode = (hashCode * 397) ^ (Criteria2 != null ? Criteria2.GetHashCode() : 0);
                return hashCode;
            }
        }
    }
}