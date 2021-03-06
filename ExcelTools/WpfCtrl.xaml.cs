﻿using System;
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
using ExcelTools.Filters;

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
            TabMan.SetTabCtrl(mainTab);
            FiltersMan.Listen();
        }





        public event PropertyChangedEventHandler PropertyChanged;

        [NotifyPropertyChangedInvocator]
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

    }
}
