using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using ExcelTools.Annotations;

namespace ExcelTools
{
    public class ValListSettings : DependencyObject
    {
        public static ValListSettings Instance { get; private set; }

        static ValListSettings() { Instance = new ValListSettings(); }

        public static readonly DependencyProperty RowWndProperty =
            DependencyProperty.Register("RowWnd", typeof(int),
                typeof(ValListSettings), new UIPropertyMetadata(30));
        public int RowWnd
        {
            get { return (int)GetValue(RowWndProperty); }
            set { SetValue(RowWndProperty, value); }
        }
        public static readonly DependencyProperty AllRowsProperty =
            DependencyProperty.Register("AllRows", typeof(bool),
                typeof(ValListSettings), new UIPropertyMetadata(false));
        public bool AllRows
        {
            get { return (bool)GetValue(AllRowsProperty); }
            set { SetValue(AllRowsProperty, value); }
        }

    }
}
