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
    /// Interaction logic for CtrlFilterBase.xaml
    /// </summary>
    public partial class CtrlFilterBase : UserControl
    {
        public CtrlFilterBase()
        {
            InitializeComponent();
        }
        private void BtnDelete_OnClick(object sender, RoutedEventArgs e)
        {
            var element = sender as FrameworkElement;
            if (element == null) return;
            var filter = element.DataContext as FilterProto;
            FilterCollection.Remove(filter);
        }

        private void Selector_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            //var items = ((ListBox)sender).SelectedItems;
            //var fiter = (object)

        }
    }
}
