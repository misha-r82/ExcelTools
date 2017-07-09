using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using ExcelTools.Annotations;
using Microsoft.Office.Interop.Excel;

namespace ExcelTools
{
    public class CellValue : IEquatable<CellValue>
    {
        public  enum CellValType {String, Numeric, Date, Time}
        public static string TIME_FORMAT = @"hh\:mm";
        public static string DATE_FORMAT { get; set; } = "dd.MM.yyyy";
        private object _xlVal;
        private string _strVal;
        private string _format;
        private bool _dateAsString;
        public CellValType Type { get; }
        public double ValDouble { get { return (double) _xlVal; } }
        public TimeSpan ValTime { get { return DateTime.FromOADate(ValDouble).TimeOfDay; } }
        public DateTime ValDate {
            get
            {
                if (!_dateAsString) return (DateTime) _xlVal;
                DateTime tmp;
                if (DateTime.TryParseExact(_xlVal.ToString(), _format, CultureInfo.InvariantCulture, 
                    DateTimeStyles.None, out tmp)) return tmp;
                return new DateTime();
            } }
                        

        public CellValue(Range rng)
        {
            _format = rng.NumberFormat.ToString();
            _xlVal = rng.Value;
            if (_xlVal == null)
            {
                if (IsNumericFormat(_format)) Type = CellValType.Numeric;
                else if (IsDateFormat(_format)) Type = CellValType.Date;
                else if (IsTimeFormat(_format)) Type = CellValType.Time;
                else Type = CellValType.String;
            }
            else
            {
                if (_xlVal is string)
                {
                    DateTime tmp;
                    if (IsDateFormat(_format) && DateTime.TryParseExact(_xlVal.ToString(), 
                        _format, CultureInfo.InvariantCulture, DateTimeStyles.None, out tmp))
                    {
                        Type = CellValType.Date;
                        _dateAsString = true;
                    }
                        
                    else Type = CellValType.String;
                }
                    
                    
                else if (_xlVal is double)
                {
                    if (IsTimeFormat(_format)) Type = CellValType.Time;
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
                if (_xlVal == null) return null;
                switch (Type)
                {
                        
                        case CellValType.String:
                        case CellValType.Numeric:
                            return _xlVal.ToString();
                        case CellValType.Date:
                            return _dateAsString ? _xlVal.ToString() : ((DateTime)_xlVal).ToString(_format);
                        case CellValType.Time:
                            double timeDouble = (double)_xlVal;
                            TimeSpan time = DateTime.FromOADate(timeDouble).TimeOfDay;
                            return time.ToString(_format);

                }
                return "";
            }
            set
            {
                if (_dateAsString)
                    _xlVal = value;
                else
                    switch (Type)
                    {
                        case CellValType.String:
                            if (_dateAsString)
                                _xlVal = value;
                            break;
                        case CellValType.Numeric:
                            double tmp = 0;
                            if (double.TryParse(value, out tmp))
                                _xlVal = tmp;
                            else _xlVal = null;
                            break;
                        case CellValType.Date:
                            if (value != null)
                            {
                                DateTime date = new DateTime();
                                if (DateTime.TryParseExact(value, _format, CultureInfo.InvariantCulture,
                                    DateTimeStyles.None, out date))
                                    _xlVal = date;
                                else _xlVal = null;
                            }
                            else _xlVal = null;
                            break;
                        case CellValType.Time:
                            if (value != null)
                            {
                                TimeSpan time = new TimeSpan();
                                if (TimeSpan.TryParseExact(value, _format, null, out time))
                                    _xlVal = new DateTime(time.Ticks).ToOADate();
                                else _xlVal = null;
                            }
                            else _xlVal = null;
                            break;
                    }

                OnPropertyChanged();
                OnPropertyChanged(nameof(XlVal));
            }
        }


        public override string ToString()
        {
            return StrVal;
        }

        public event PropertyChangedEventHandler PropertyChanged;

        [NotifyPropertyChangedInvocator]
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public bool Equals(CellValue other)
        {
            if (ReferenceEquals(null, other)) return false;
            if (ReferenceEquals(this, other)) return true;
            return Equals(_xlVal, other._xlVal) && Type == other.Type;
        }

        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj)) return false;
            if (ReferenceEquals(this, obj)) return true;
            if (obj.GetType() != this.GetType()) return false;
            return Equals((CellValue) obj);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                return ((_xlVal != null ? _xlVal.GetHashCode() : 0) * 397) ^ (int) Type;
            }

        }
    }
}
