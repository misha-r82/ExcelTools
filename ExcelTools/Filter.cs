using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls.Primitives;
using ExcelTools.Annotations;
using Microsoft.Office.Interop.Excel;
using System.Windows;

namespace ExcelTools
{
    public abstract class FilterProto : INotifyPropertyChanged
    {
        private bool _enabled;

        public FilterProto()
        {
            FilterRng = Current.CurRegion.ActiveCell;
            int col = FilterRng.Column - Current.CurRegion.firstCol;
            Name = Current.CurRegion.ActiveRow.Cells[col].ColName;
            ColNum =  col + 1;
            Enabled = true;
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

        private int ColNum { get; }
        public string Name { get; }
        
        private Range FilterRng { get; set; }
        protected virtual object Criteria1 { get; }
        protected virtual object Criteria2 { get; }

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

    class StrFilter : FilterProto
    {
        private string _patt;

        public string Patt
        {
            get { return _patt; }
            set
            {
                if(_patt == value) return;
                _patt = value;
                OnPropertyChanged();
                SetFilter();
            }
        }

        protected override object Criteria1
        {
            get { return string.Format("*{0}*", Patt); }
        }

        public StrFilter(Cell cell) : base()
        {
            _patt = "";
        }
    }

    class NumericFilter : FilterProto
    {
        public double From { get; set; }
        public double To { get; set; }

        protected override object Criteria1
        {
            get { return string.Format(">={0}", From); }
        }
        protected override object Criteria2
        {
            get { return string.Format("<={0}", To); }
        }
        public NumericFilter(Cell cell) : base()
        {
            var values = cell.ValList.OfType<double>().ToArray();
            if (values.Any())
            {
                From = values.Min();
                To = values.Max();
            }
            else From = To = 0;
        }
    }
    class DateFilter : FilterProto
    {
        private DateTime _from;
        private DateTime _to;

        public DateFilter(Cell cell) : base()
        {
            {
                From = DateTime.Now.AddDays(-1);
                To = DateTime.Now;
            }
        }

        public DateTime From
        {
            get { return _from; }
            set
            {
                if (value == _from) return;
                _from = value;
                OnPropertyChanged();
            }
        }

        public DateTime To
        {
            get { return _to; }
            set
            {
                if (value == _to) return;
                _to = value;
                OnPropertyChanged();
            }
        }

        protected override object Criteria1
        {
            get { return string.Format(">={0}", From.ToString(@"MM\/dd\/yyyy")); }
        }
        protected override object Criteria2
        {
            get { return string.Format("<={0}", To.ToString(@"MM\/dd\/yyyy")); }
        }

    }

    class TimeFilter : FilterProto
    {
        public TimeSpan From { get; set; }
        public TimeSpan To { get; set; }
        public TimeSpan[] ValList { get; private set; }
        protected override object Criteria1
        {
            get { return string.Format(">={0}", From); }
        }
        protected override object Criteria2
        {
            get { return string.Format("<={0}", To); }
        }
        public TimeFilter(Cell cell) : base()
        {
            ValList = cell.ValList.OfType<TimeSpan>().ToArray();
            if (ValList.Any())
            {
                From = ValList.Min();
                To = ValList.Max();                
            }
            else
            {
                From = new TimeSpan(0);
                To = new TimeSpan(23,59,59);
            }

        }        
    }
}
