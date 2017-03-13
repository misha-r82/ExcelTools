using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace WsManager
{
    public class WbMan : IEnumerable<ExWb>, INotifyCollectionChanged
    {
        public WbMan()
        {
            _wbList = new List<ExWb>();
            WsList = new ObservableCollection<ExWs>();
        }

        public void OpenWbs(bool write = false)
        {
            foreach (ExWb exWb in _wbList)
                exWb.Open(true);
        }
        public void CloseWbs()
        {
            foreach (ExWb exWb in _wbList)
                exWb.Close();
        }
        private List<ExWb> _wbList { get; }
        public  ObservableCollection<ExWs> WsList { get; }


        public  void LoadFiles(IEnumerable<string> files)
        {
            foreach (string file in files)
                if (!_wbList.Any(f => f.File == file))
                {
                    var wb = new ExWb(file);
                    _wbList.Add(wb);
                    foreach (var ws in wb.WorkSheets)
                        WsList.Add(ws);
                }
            OnListChanged();                       
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public IEnumerator<ExWb> GetEnumerator()
        {
            return _wbList.GetEnumerator();
        }

        public event NotifyCollectionChangedEventHandler CollectionChanged;
        protected virtual void OnListChanged()
        {
            CollectionChanged?.Invoke(this, new NotifyCollectionChangedEventArgs(NotifyCollectionChangedAction.Reset));
        }
    }
}
