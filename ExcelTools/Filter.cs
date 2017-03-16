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
            ColNum = Current.CurRegion.firstCol;
            /*Range region = cells.CurrentRegion;
            int left = ((Range)region.Cells[0,0]).Column;
            ColNum = cells.Column - left;
            if (cells.Rows.Count == 1)
                FilterRng = region;
            else
            {
                int right = left + region.Columns.Count;
                int top = ((Range) cells[0, 0]).Row;
                int bottom = top + cells.Rows.Count;
                Worksheet ws = cells.Worksheet;
                FilterRng = ws.get_Range(ws.Cells[top, left], ws.Cells[bottom, right]);
            }*/
        }

        public bool Enabled
        {
            get { return _enabled; }
            set
            {
                if (_enabled == value) return;
                _enabled = value;
                OnPropertyChanged();
                if (_enabled) SetFilter();
                else RemoveFilter();
            }
        }

        private int ColNum { get; }
        public string Name { get; }
        
        private Range FilterRng { get; set; }
        protected virtual object Criteria { get; }

        public void SetFilter()
        {
            FilterRng.AutoFilter(ColNum, Criteria);
        }
        public void RemoveFilter()
        {
            FilterRng.AutoFilter(ColNum, "");
        }
        public event PropertyChangedEventHandler PropertyChanged;

        [NotifyPropertyChangedInvocator]
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
    class StrFilter : FilterProto
    {
        public string Patt { get; set; }
        protected override object Criteria
        {
            get { return string.Format("*{0}*", Patt); }
        }

        public StrFilter(Range cells) : base()
        {
            Patt = "1";
            //SetFilter();
        }
    }

    class NumericFilter : FilterProto
    {
        public int From { get; set; }
        public int To { get; set; }

        protected override object Criteria
        {
            get { return new string[] {string.Format(">={0}", From), string.Format("<={0}", To)}; }
        }

        public NumericFilter(Range cells) : base()
        {
            From = 23;
            To = 27;
            //SetFilter();
        }
    }
    class DateFilter : FilterProto
    {
        public DateFilter(Range cells) : base()
        {
            {
                From = DateTime.Now.AddDays(-1);
                To = DateTime.Now;
            }
        }
        public DateTime From { get; set; }
        public DateTime To { get; set; }
        protected override object Criteria
        {
            get { return new string[] { string.Format(">={0}", From), string.Format("<={0}", To) }; }
        }

   
    }

    class TimeFilter : FilterProto
    {
        public TimeSpan From { get; set; }
        public TimeSpan To { get; set; }
        protected override object Criteria
        {
            get { return new string[] { string.Format(">={0}", From), string.Format("<={0}", To) }; }
        }

        public TimeFilter(Range cells) : base()
        {
            From = new TimeSpan(0);
            To = new TimeSpan(0,23,59,59);
        }        
    }
}
