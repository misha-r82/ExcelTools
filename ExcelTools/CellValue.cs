using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using ExcelTools.Annotations;
using Microsoft.Office.Interop.Excel;

namespace ExcelTools
{
    class CellValue : INotifyPropertyChanged
    {
        public  enum CellValType {String, Numeric, Date, Time}
        public static string TIME_FORMAT = "00:00";
        public static string DATE_FORMAT = "";
        private object _xlVal;
        public CellValType Type { get; }

        public CellValue(Range rng)
        {
            string format = rng.NumberFormat.ToString();
            _xlVal = rng.Value;
            if (_xlVal == null)
            {
                if (IsNumericFormat(format)) Type = CellValType.Numeric;
                else if (IsDateFormat(format)) Type = CellValType.Date;
                else if (IsTimeFormat(format)) Type = CellValType.Time;
                else Type = CellValType.String;
            }
            else
            {
                if (_xlVal is string)
                    Type = CellValType.String;
                else if (_xlVal is double)
                {
                    if (IsTimeFormat(format)) Type = CellValType.Time;
                    else Type = CellValType.Numeric;
                }
                else if (_xlVal is DateTime)
                    Type = CellValType.Date;
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

        public object XlVal
        {
            get { return _xlVal; }
            set
            {
                if (Equals(value, _xlVal)) return;
                _xlVal = value;
                OnPropertyChanged();
            }
        }

        public string StrVal
        {
            get
            {
                switch (Type)
                {
                        case CellValType.String:
                        case CellValType.Numeric:
                            return _xlVal.ToString();
                        //case CellValType.Date:
                }
                return "";
            }
        }
        public event PropertyChangedEventHandler PropertyChanged;

        [NotifyPropertyChangedInvocator]
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
