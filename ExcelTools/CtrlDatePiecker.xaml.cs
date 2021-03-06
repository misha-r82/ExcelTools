﻿using System;
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

namespace DatePiecker
{
    /// <summary>
    /// Interaction logic for CtrlDatePiecker.xaml
    /// </summary>
    public partial class CtrlDatePiecker : UserControl
    {
        public CtrlDatePiecker()
        {
            InitializeComponent();

        }
        public static readonly DependencyProperty DateProperty;
        static CtrlDatePiecker()
        {           
            DateProperty = DependencyProperty.Register(
            "Date", typeof(DateTime),
            typeof(CtrlDatePiecker), new FrameworkPropertyMetadata
            { BindsTwoWayByDefault = true, });
        }
        public DateTime Date
        {
            get { return (DateTime)GetValue(DateProperty); }
            set { SetValue(DateProperty, value); }
        }



        private void datepicker_Loaded(object sender, RoutedEventArgs e)
        {
            Button button = (Button)datepicker.Template.FindName("PART_Button", datepicker);
            button.Width = Double.NaN;
            button.Height = Double.NaN;
            button.Template = (ControlTemplate)FindResource("btnTemplate");
        }
    }
}
