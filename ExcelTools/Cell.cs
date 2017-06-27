using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Runtime.CompilerServices;
using ExcelTools.Annotations;
using Microsoft.Office.Interop.Excel;

namespace ExcelTools
{
    public class Cell : INotifyPropertyChanged
    {
        private string _colName;
        private string _strVal;
        private double _numVal;
        private TimeSpan _timeVal;
        private DateTime _dateVal;
        private Range _rng;
        private bool _isSelected;

        public CellTypes Type { get; }
        public Cell[] ValList { get; private set; }
        public Range Rng { get { return _rng; } }
        public string ColName
        {
            get { return _colName; }
            set { _colName = value; }
        }
        public double NumVal
        {
            get { return _numVal; }
            set
            {
                _numVal = value;
                _rng.Value = value;
                OnPropertyChanged(nameof(IsValid));
            }
        }

        public DateTime DateVal
        {
            get { return _dateVal; }
            set
            {
                _dateVal = value;
                _rng.Value = value;
                OnPropertyChanged(nameof(IsValid));
            }
        }

        public TimeSpan TimeVal
        {
            get { return _timeVal; }
            set
            {
                _timeVal = value;
                var date = new DateTime(0) + value;
                OnPropertyChanged(nameof(IsValid));
                _rng.Value = date.ToOADate();
            }
        }

        public bool IsValid
        {
            get
            {
                if (_rng.Validation.Value) return true;
                return false;
            }
        }
        public string StrVal
        {
            get { return _strVal; }
            set
            {
                _strVal = value;
                _rng.Value = value;
                OnPropertyChanged(nameof(IsValid));
            }
        }
        public bool IsSelected
        {
            get { return _isSelected; }
            set
            {
                if (value == IsSelected) return;
                _isSelected = value;
                OnPropertyChanged();
            }
        }

        private bool IsTimeFormat(string cellFormat)
        {
            return cellFormat.Contains('h');
        }

        private bool IsNumericFormat(string cellFormat)
        {
            if (cellFormat.Contains('0')) return true;
            if (cellFormat.Contains('#')) return true;
            if (cellFormat.Contains("%")) return true;
            return false;
        }
        private bool IsDateFormat(string cellFormat)
        {
            if (IsTimeFormat(cellFormat)) return false;
            return cellFormat.Contains('m') || cellFormat.Contains('y') || cellFormat.Contains('d');
        }
        public Cell(Range rng, bool setValList)
        {
            _rng = rng;
            string format = rng.NumberFormat.ToString();
            object val = rng.Value;
            if (val == null)
            {
                if (IsNumericFormat(format)) Type = CellTypes.numeric;
                else if(IsDateFormat(format)) Type = CellTypes.date;
                else if(IsTimeFormat(format)) Type = CellTypes.time;
                else Type = CellTypes.str;
            }
            else
            {
                if (val is string)
                {
                    Type = CellTypes.str;
                    _strVal = (string) val;
                }
                else if (val is double)
                {
                    if (IsTimeFormat(format))
                    {
                        Type = CellTypes.time;
                        _timeVal = DateTime.FromOADate((double) val).TimeOfDay;
                    }
                    else
                    {
                        Type = CellTypes.numeric;
                        _numVal = (double) val;
                    }
                } else if (val is DateTime)
                {
                    Type = CellTypes.date;
                    _dateVal = (DateTime) val;
                }
            }          
            var cr = Current.CurRegion;
            Range colNameRng = (Range)cr.ActiveWs.Cells[cr.firstRow - 1, rng.Column];
            _colName = colNameRng.Value != null ? colNameRng.Value.ToString() : "";
            if (!setValList) return;
            int row = rng.Row;

            int from = row - cr.Wnd > cr.firstRow ? row - cr.Wnd : cr.firstRow;
            int to = row + cr.Wnd < cr.lastRow ? row + cr.Wnd : cr.lastRow;
            var ws = rng.Worksheet;
            Range col = ws.Range[ws.Cells[from, rng.Column], ws.Cells[to, rng.Column]];
            var tmp = new List<Cell>();
            foreach (object r in col)
                tmp.Add(new Cell((Range) r, false));
            ValList = tmp
                .Where(c=>c.Type == Type && !string.IsNullOrEmpty(c.ToString()))
                .Distinct().ToArray();
        }

        
        public override string ToString()
        {
            switch (Type)
            {
                case CellTypes.str: return StrVal;
                case CellTypes.numeric: return NumVal.ToString(CultureInfo.InvariantCulture);
                case CellTypes.date: return DateVal.ToString(CultureInfo.InvariantCulture);
                case CellTypes.time: return TimeVal.ToString();
                default: return "";
            }
        }

        public object Value
        {
            get
            {
                switch (Type)
                {
                    case CellTypes.str: return StrVal;
                    case CellTypes.numeric: return NumVal;
                    case CellTypes.date: return DateVal;
                    case CellTypes.time: return TimeVal;
                    default: return "";
                }
            }
        }


        public event PropertyChangedEventHandler PropertyChanged;

        [NotifyPropertyChangedInvocator]
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    public class CellValEquilityComparer : IEqualityComparer<Cell>
    {
        public bool Equals(Cell x, Cell y)
        {
            return x.Value.ToString() == y.Value.ToString();
        }

        public int GetHashCode(Cell obj)
        {
            return obj.Value.ToString().GetHashCode();
        }
    }
}