using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls.Primitives;
using Microsoft.Office.Interop.Excel;

namespace ExcelTools
{
    public abstract class FilterProto
    {
        public FilterProto(Range cells)
        {
            FilterRng = Current.CurRegion.CurRng;
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
        public bool Enabled { get; set; }
        private int ColNum { get; }
        public string Name { get
        { 
            var cell = (Range) FilterRng.Cells[1, 1];
            if (cell.Value2 == null) return "";
            return cell.Value2.ToString();
        }
        }
        private Range FilterRng { get; set; }
        protected virtual object Criteria { get; }

        public void SetFilter()
        {
            FilterRng.AutoFilter(ColNum, Criteria);
        }

    }
    class StrFilter : FilterProto
    {
        public string Patt { get; set; }
        protected override object Criteria
        {
            get { return string.Format("*{0}*", Patt); }
        }

        public StrFilter(Range cells) : base(cells)
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

        public NumericFilter(Range cells) : base(cells)
        {
            From = 23;
            To = 27;
            //SetFilter();
        }
    }
    class DateFilter : FilterProto
    {
        public DateFilter(Range cells) : base(cells)
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

        public TimeFilter(Range cells) : base(cells)
        {
            From = new TimeSpan(0);
            To = new TimeSpan(0,23,59,59);
        }        
    }
}
