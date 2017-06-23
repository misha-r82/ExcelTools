using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
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
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Office.Interop.Excel;

namespace ExcelTools
{
    /// <summary>
    /// Interaction logic for CtrlFilters.xaml
    /// </summary>
    public partial class CtrlCalk : UserControl
    {
        public CtrlCalk()
        {
            InitializeComponent();

            Current.CurRegion.PropertyChanged += (sender, args) =>
            {
                object val = Current.CurRegion.ActiveCell.Value;
                if (val is double || val is long)
                {
                    ctrlCalk.DisplayText = Current.CurRegion.ActiveCell.Value2.ToString();
                }
            };
        }

    }
}
