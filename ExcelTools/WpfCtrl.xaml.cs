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
        public ObservableCollection<FilterProto> Filters { get; set; }


        public WpfCtrl()
        {
            InitializeComponent();
            Filters = new ObservableCollection<FilterProto>();
            DataContext = this;
            Current.CurRegion.PropertyChanged += (sender, args) =>
            { if (args.PropertyName == "CurRegion.Selection") lstActiveRow.Items.Refresh();};


        }

        private void button_Click(object sender, RoutedEventArgs e)
        {
            var flt = FilterFactory.CreateFilter();
            if (flt != null) Filters.Add(flt);
            txtText.Text+= ((Range) ThisWorkbook.app.Selection).Value.GetType().ToString() + "\n";
            //Filters.Add(flt);
            //

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
                case "btnNewRow": curReg.CurRowNumInRng = curReg.CurRng.Rows.Count; break;
            }
        }

        private void TimePicker_OnDataContextChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            var timePicker = sender as Xceed.Wpf.Toolkit.TimePicker;
            var tmp = UIHelper.GetChildOfType<ComboBox>(timePicker);
        }

        private void TimePicker_OnInitialized(object sender, EventArgs e)
        {
            var timePicker = sender as Xceed.Wpf.Toolkit.TimePicker;
            FieldInfo fieldInfo = typeof(Xceed.Wpf.Toolkit.TimePicker).GetField("_timeListBox", BindingFlags.Instance | BindingFlags.NonPublic);
            var tb = (System.Windows.Controls.ListBox)fieldInfo.GetValue(timePicker);

        }
    }
}
