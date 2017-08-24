using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTools.Filters
{
    public class FiltersMan
    {
        private static string _prewRangeAddr;
        private static string _prewSheetName;

        public static void Listen()
        {
            FilterCollection.Filters.CollectionChanged += FiltersOnCollectionChanged;
            Current.CurRegion.PropertyChanged += CurRegionOnPropertyChanged;
            Current.CurRegion.ActiveWs.Application.SheetDeactivate += Application_SheetDeactivate;
        }

        private static void Application_SheetDeactivate(object Sh)
        {
            foreach (var filter in FilterCollection.Filters)
                    filter.RemoveFilter();
        }

        private static void CurRegionOnPropertyChanged(object sender, PropertyChangedEventArgs propertyChangedEventArgs)
        {
            if (propertyChangedEventArgs.PropertyName != "ActiveRow") return;
            if (!string.IsNullOrEmpty(_prewRangeAddr) && _prewRangeAddr != Current.CurRegion.TblRange.CurRng.Address)
                foreach (var filter in FilterCollection.Filters)
                    filter.OnRangeChange();
            

            _prewRangeAddr = Current.CurRegion.TblRange.CurRng.Address;
            _prewSheetName = Current.CurRegion.ActiveWs.Name;


        }

        private static void FiltersOnCollectionChanged(object sender, NotifyCollectionChangedEventArgs notifyCollectionChangedEventArgs)
        {
            if (notifyCollectionChangedEventArgs.Action == NotifyCollectionChangedAction.Add)
                foreach (var filter in FilterCollection.Filters)
                    filter.SetFilter();
        }

    }
}
