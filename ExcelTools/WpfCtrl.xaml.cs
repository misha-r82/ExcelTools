using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
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
using ExcelTools.Annotations;
using Microsoft.Office.Interop.Excel;
using static ExcelTools.Current;
using Button = System.Windows.Controls.Button;


namespace ExcelTools
{
    /// <summary>
    /// Interaction logic for WpfCtrl.xaml
    /// </summary>
    public partial class WpfCtrl : UserControl, INotifyPropertyChanged
    {
        public WpfCtrl()
        {
            InitializeComponent();
            DataContext = this;
            Current.CurRegion.PropertyChanged += (sender, args) =>
            { if (args.PropertyName == "CurRegion.Selection") lstActiveRow.Items.Refresh();};
        }



        private void TabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }



        public event PropertyChangedEventHandler PropertyChanged;

        [NotifyPropertyChangedInvocator]
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        private void BtnFrist_OnClick(object sender, RoutedEventArgs e)
        {
            var btn = sender as Button;
            var curReg = Current.CurRegion;
            switch (btn.Name)
            {
                case "btnFrist": curReg.CurRowNumInRng = 1; break;
                case "btnLast": curReg.CurRowNumInRng = curReg.CurRng.Rows.Count; break;
                case "btnNext": if (curReg.CurRowNumInRng < curReg.CurRng.Rows.Count -1) curReg.CurRowNumInRng++ ; break;
                case "btnPrev": if(curReg.CurRowNumInRng > 1) curReg.CurRowNumInRng--; break;
                case "btnNewRow": 
                    var oldRow = curReg.ActiveRow;
                    curReg.CurRowNumInRng = curReg.CurRng.Rows.Count;
                    for (int i = 0; i < oldRow.Cells.Length; i++)
                        if (oldRow.Cells[i].IsSelected)
                            curReg.ActiveRow.Cells[i].Rng.Value2 = oldRow.Cells[i].Rng.Value2;
                    break;
            }
        }


        private void ChkAll_OnClick(object sender, RoutedEventArgs e)
        {
            var val = chkAll.IsChecked == true;
            foreach (Cell cell in Current.CurRegion.ActiveRow.Cells)
                cell.IsSelected = val;
        }

    }
}
