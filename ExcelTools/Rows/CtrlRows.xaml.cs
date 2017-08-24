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
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace ExcelTools
{
    /// <summary>
    /// Interaction logic for CtrlRows.xaml
    /// </summary>
    public partial class CtrlRows : UserControl
    {
        public CtrlRows()
        {
            InitializeComponent();
            Current.CurRegion.PropertyChanged += (sender, args) =>
            { if (args.PropertyName == "CurRegion.Selection") lstActiveRow.Items.Refresh(); };
        }

        private void BtnFrist_OnClick(object sender, RoutedEventArgs e)
        {

            var btn = sender as Button;
            var curReg = Current.CurRegion;
            switch (btn.Name)
            {
                case "btnFrist":
                    curReg.CurRowNumInRng = 1;
                    break;
                case "btnLast":
                    curReg.CurRowNumInRng = curReg.TblRange.CurRng.Rows.Count;
                    break;
                case "btnNext":
                    if (curReg.CurRowNumInRng < curReg.TblRange.CurRng.Rows.Count - 1) curReg.CurRowNumInRng++;
                    break;
                case "btnPrev":
                    if (curReg.CurRowNumInRng > 1) curReg.CurRowNumInRng--;
                    break;
                case "btnNewRow":
                    var oldRow = curReg.ActiveRow;
                    curReg.CurRowNumInRng = curReg.TblRange.CurRng.Rows.Count;
                    for (int i = 0; i < oldRow.ExCells.Length; i++)
                        if (oldRow.ExCells[i].IsSelected)
                            curReg.ActiveRow.ExCells[i].Rng.Value2 = oldRow.ExCells[i].Rng.Value2;
                    break;
            }
        }


        private void ChkAll_OnClick(object sender, RoutedEventArgs e)
        {
            var val = chkAll.IsChecked == true;
            foreach (ExCell cell in Current.CurRegion.ActiveRow.ExCells)
                cell.IsSelected = val;
        }

    }
}
