﻿using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls.Primitives;

namespace ExcelTools
{
    class StrFilter : FilterProto
    {
        private string _patt;
        public override string Name { get { return base.Name + " содержит"; } }

        public string Patt
        {
            get { return _patt; }
            set
            {
                if(_patt == value) return;
                _patt = value;
                OnPropertyChanged();                
            }
        }

        protected override object Criteria1
        {
            get { return string.Format("*{0}*", Patt); }
        }

        public StrFilter(ExCell exCell) : base()
        {
            _patt = "";
        }
    }

    class NumericFilter : FilterProto
    {
        private double _from;
        private double _to;

        public double From
        {
            get { return _from; }
            set
            {
                if (value == _from) return;
                _from = value;
                OnPropertyChanged();
            }
        }

        public double To
        {
            get { return _to; }
            set
            {
                if (value == _to) return;
                _to = value;
                OnPropertyChanged();
            }
        }

        public override string Name { get { return base.Name + " между"; } }
        protected override object Criteria1
        {
            get { return string.Format(">={0}", From); }
        }
        protected override object Criteria2
        {
            get { return string.Format("<={0}", To); }
        }
        public NumericFilter(ExCell exCell) : base()
        {
            var values = exCell.ValList.Where(v => v.Type == CellValue.CellValType.Numeric).
                Select(v => v.ValDouble).ToArray();
            if (values.Any())
            {
                _from = values.Min();
                _to = values.Max();
            }
            else From = _from = _to;
        }
    }
    class DateFilter : FilterProto
    {
        private DateTime _from;
        private DateTime _to;
        public override string Name { get { return base.Name + " между"; } }
        public DateFilter(ExCell exCell) : base()
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
        private TimeSpan _from;
        private TimeSpan _to;

        public TimeSpan From
        {
            get { return _from; }
            set
            {
                if (value == _from) return;
                _from = value;
                OnPropertyChanged();
            }
        }

        public TimeSpan To
        {
            get { return _to; }
            set
            {
                if (value == _to) return;
                _to = value;
                OnPropertyChanged();
            }
        }

        public override string Name { get { return base.Name + " между"; } }
        public TimeSpan[] ValList { get; private set; }
        protected override object Criteria1
        {
            get { return string.Format(">={0}", From); }
        }
        protected override object Criteria2
        {
            get { return string.Format("<={0}", To); }
        }
        public TimeFilter(ExCell exCell) : base()
        {
            ValList = exCell.ValList.Where(v=>v.Type == CellValue.CellValType.Time).
                Select(v=>v.ValTime).OrderBy(v=>v).ToArray();
            if (ValList.Any())
            {
                _from = ValList.Min();
                _to = ValList.Max();                
            }
            else
            {
                _from = new TimeSpan(0);
                _to = new TimeSpan(23,59,59);
            }

        }        
    }
}
