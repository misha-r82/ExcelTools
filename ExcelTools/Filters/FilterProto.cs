using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows;
using ExcelTools.Annotations;
using Microsoft.Office.Interop.Excel;

namespace ExcelTools
{
    public abstract class FilterProto : INotifyPropertyChanged
    {
        private int ColNum { get; }
        public string Name { get; }
        public object[] ValueList { get; set; }
        private Range FilterRng { get; set; }
        protected virtual object Criteria1 { get; }
        protected virtual object Criteria2 { get; }
        private bool _enabled;
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
            Name = Current.CurRegion.ActiveRow.Cells[col].ColName;
            var tmpCell = new Cell(FilterRng, true);
            ValueList = tmpCell.ValList;
            ColNum =  col + 1;
            Enabled = true;
        }
        public void SetFilter()
        {
            if (Enabled)
            {
                try
                {
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
            FilterRng.AutoFilter(ColNum);
        }
        public event PropertyChangedEventHandler PropertyChanged;

        [NotifyPropertyChangedInvocator]
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            SetFilter();
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}