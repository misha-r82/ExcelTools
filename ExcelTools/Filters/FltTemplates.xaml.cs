using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace ExcelTools
{
    /// <summary>
    /// Interaction logic for FrmTest.xaml
    /// </summary>
    public partial class FltTemplates
    {
        private void BtnTimeFltReset_OnClick(object sender, RoutedEventArgs e)
        {
            var btn = sender as Button;
            if (btn == null) return;
            var flt = btn.DataContext as DateFilter;
            if (flt == null) return;
            DateTime now = DateTime.Now.Date;
            int dayOfWeek = (int)now.DayOfWeek;
            int quart = (now.Month - 1) / 3 + 1;
            int halfYear = (now.Month - 1) / 6 + 1;
            dayOfWeek = dayOfWeek == 0 ? 6 : dayOfWeek - 1; // c нуля
            switch (btn.Name)
            {
                case "btnToday": flt.From = flt.To = now; break;
                case "btnYestarday": flt.From = flt.To = now.AddDays(-1); break;
                case "btnCurWeek":
                    flt.From = now.AddDays(-dayOfWeek);
                    flt.To = now;
                    break;
                case "btnLastWeek":
                    flt.From = now.AddDays(-dayOfWeek - 7);
                    flt.To = now.Date.AddDays(-dayOfWeek - 1);
                    break;
                case "btnCurMonth":
                    flt.From = now.Date.AddDays(-now.Day + 1);
                    flt.To = now;
                    break;
                case "btnLastMonth":
                    flt.To = now.Date.AddDays(-now.Day);
                    flt.From = now.Date.AddDays(-now.Day - flt.To.AddDays(-1).Day);

                    break;
                case "btnCurQuart":
                    flt.To = now;
                    flt.From = new DateTime(now.Year, 1, 1).AddMonths(3 * (quart - 1));

                    break;
                case "btnLastQuart":
                    flt.From = new DateTime(now.Year, 1, 1).AddMonths(3 * (quart - 2));
                    flt.To = new DateTime(now.Year, 1, 1).AddMonths(3 * (quart - 1)).AddDays(-1);
                    break;
                case "btnCurHalfYear":
                    flt.To = now;
                    flt.From = new DateTime(now.Year, 1, 1).AddMonths(3 * (halfYear - 1));

                    break;
                case "btnLastHalfYear":
                    flt.To = new DateTime(now.Year, 1, 1).AddMonths(3 * (halfYear - 1)); ;
                    flt.From = new DateTime(now.Year, 1, 1).AddMonths(3 * (halfYear - 2)); ;

                    break;
                case "btnCurYear":
                    flt.From = new DateTime(now.Year, 1, 1);
                    flt.To = now;
                    break;
                case "btnLastYear":
                    flt.From = new DateTime(now.Year - 1, 1, 1);
                    flt.To = new DateTime(now.Year, 1, 1).AddDays(-1);
                    break;
                case "btn7days":
                    flt.From = now.AddDays(-7);
                    flt.To = now;
                    break;
                case "btn14days":
                    flt.From = now.AddDays(-14);
                    flt.To = now;
                    break;
                case "btn30days":
                    flt.From = now.AddDays(-30);
                    flt.To = now;
                    break;
                case "btn60days":
                    flt.From = now.AddDays(-60);
                    flt.To = now;
                    break;
                case "btn90days":
                    flt.From = now.AddDays(-90);
                    flt.To = now;
                    break;
                case "btn180days":
                    flt.From = now.AddDays(-180);
                    flt.To = now;
                    break;
                case "btn360days":
                    flt.From = now.AddDays(-360);
                    flt.To = now;
                    break;

            }

        }
    }
}
