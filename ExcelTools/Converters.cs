using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Data;

namespace ExcelTools.Converters
{
    [ValueConversion(typeof(DateTime), typeof(TimeSpan))]
    public class TimeToDateConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter,
            System.Globalization.CultureInfo culture)
        {
            if (value == null) return "00:00";
            if (value is ExCell) value = ((ExCell) value).Value;
            return  ((TimeSpan)value).ToString(@"hh\:mm");
        }

        public object ConvertBack(object value, Type targetType, object parameter,
            System.Globalization.CultureInfo culture)
        {
            TimeSpan rez;
            if (value == null) return new TimeSpan();
            if (TimeSpan.TryParse(value.ToString(), out rez))
                return rez;
            return new TimeSpan();

        }
    }
    public class InvBoolConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter,
            System.Globalization.CultureInfo culture)
        {
            bool val = value != null && (bool)value;
            return !val;
        }

        public object ConvertBack(object value, Type targetType, object parameter,
            System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();

        }
    }
    public class BoolToVisConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter,
            System.Globalization.CultureInfo culture)
        {
            bool val = value != null && (bool) value;
            return val ? Visibility.Visible : Visibility.Collapsed;
        }

        public object ConvertBack(object value, Type targetType, object parameter,
            System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();

        }
    }
    public class InvBoolToVisConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter,
            System.Globalization.CultureInfo culture)
        {
            bool val = value != null && (bool)value;
            return val ?  Visibility.Collapsed : Visibility.Visible;
        }

        public object ConvertBack(object value, Type targetType, object parameter,
            System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();

        }
    }
    public class TestConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter,
            System.Globalization.CultureInfo culture)
        {
            return value;
        }

        public object ConvertBack(object value, Type targetType, object parameter,
            System.Globalization.CultureInfo culture)
        {
            return value;

        }
    }
}
