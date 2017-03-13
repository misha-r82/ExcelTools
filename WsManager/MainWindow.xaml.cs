using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
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
using Window = System.Windows.Window;
using Excel = Microsoft.Office.Interop.Excel;
namespace WsManager
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window, INotifyPropertyChanged
    {

        public MainWindow()
        {
            InitializeComponent();
            ctrlInWb.DataContext = Data.InWb;
            ctrlOutWb.DataContext = Data.OutWb;
            dgWs.DataContext = Data.InWb;
        }
        

        public event PropertyChangedEventHandler PropertyChanged;

        [NotifyPropertyChangedInvocator]
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        private void CtrlInWb_OnPropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            //WbMan.LoadInFiles(ctrlInWb.Files);
        }

        private void LstWs_OnMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            var tmp = ((FrameworkElement)sender).DataContext;
        }

        private void ChkAll_OnClick(object sender, RoutedEventArgs e)
        {
            var chk = (CheckBox) sender;
            var items = dgWs.SelectedItems.OfType<ExWs>();
            foreach (ExWs ws in items)
                ws.IsSelected = chk.IsChecked == true;
        }

        private void BtnCopy_OnClick(object sender, RoutedEventArgs e)
        {
            foreach (ExWb wb in Data.OutWb)
            {
                wb.Open(true);
                Data.InWb.OpenWbs(true);
                foreach (ExWs ws in Data.InWb.WsList)
                    wb.Add(ws, ws.Name);
                Data.InWb.CloseWbs();
                wb.Close();
            }
        }
    }
}
