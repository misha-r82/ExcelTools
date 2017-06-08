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
using System.Threading;

namespace ExcelTools
{
    /// <summary>
    /// Interaction logic for CtrlFilters.xaml
    /// </summary>
    public partial class CtrlFilters : UserControl
    {
        public CtrlFilters()
        {
            InitializeComponent();
            Filters = new ObservableCollection<FilterProto>();
            DataContext = this;
        }
        public ObservableCollection<FilterProto> Filters { get; }
        private void button_Click(object sender, RoutedEventArgs e)
        {
            var flt = FilterFactory.CreateFilter();
            if (flt != null) Filters.Add(flt);
            //txtText.Text += ((Range)ThisWorkbook.app.Selection).Value.GetType().ToString() + "\n";
            //Filters.Add(flt);
            //

        }

        private void ListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (e.AddedItems.Count == 0) return;
            var item = (ListBoxItem)e.AddedItems[0];
            Thread.Sleep(500);
            item.IsSelected = false;
        }
    }
}
