using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using ExcelTools;

namespace ExcelTools
{
    public static class TabMan
    {
        private static TabControl _tabCtrl;

        public static void SetTabCtrl(TabControl tabCtrl)
        {
            _tabCtrl = tabCtrl;
            _tabCtrl.SelectionChanged += TabCtrlOnSelectionChanged;
        }

        private static void TabCtrlOnSelectionChanged(object sender, SelectionChangedEventArgs args)
        {
            
        }

        public static TabItem CurTab
        {
            get { return _tabCtrl.SelectedItem as TabItem; }
        }
    }
}
