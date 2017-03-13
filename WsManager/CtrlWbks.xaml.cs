using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
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
using Microsoft.Win32;
using WsManager.Annotations;

namespace WsManager
{
    /// <summary>
    /// Interaction logic for CtrlWbks.xaml
    /// </summary>
    public partial class CtrlWbks : UserControl
    {

        public CtrlWbks()
        {
            InitializeComponent();
        }

        private WbMan WbMan
        {
            get { return DataContext as WbMan; }
        }
        private void BtnOpen_OnClick(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog();
            dialog.Multiselect = true;
            dialog.Filter = "ExcelWorkBook (*.xlsx)|*.xlsx|2003 Excel Workbook (*.xls)|*.xls";
            if (dialog.ShowDialog() != true) return;
            WbMan.LoadFiles(dialog.FileNames);

        }
    }
}
