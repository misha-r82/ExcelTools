using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using ExcelTools.Annotations;
using Microsoft.Office.Interop.Excel;

namespace ExcelTools
{
    public class ValListGenerator : INotifyPropertyChanged
    {
        private int _wnd;
        private bool _allRows;

        static ValListGenerator()
        {
            Instance = new ValListGenerator();
            Instance.Wnd = 50;
        }
        public static ValListGenerator Instance { get; }
        public int Wnd
        {
            get { return _wnd; }
            set
            {
                if (value == _wnd) return;
                _wnd = value;
                OnPropertyChanged();
            }
        }

        public bool AllRows
        {
            get { return _allRows; }
            set
            {
                if (value == _allRows) return;
                _allRows = value;
                OnPropertyChanged();
            }
        }

        public static CellValue[] GetValList(Range cell, CellValue.CellValType type)
        {
            var tblRng = new TblRange(cell);
            int row = cell.Row;
            int wnd = ValListGenerator.Instance.Wnd;
            int from, to;
            if (ValListGenerator.Instance.AllRows)
            {
                from = tblRng.FirstRow;
                to = tblRng.LastRow;
            }
            else
            {
                from = row - wnd > tblRng.FirstRow ? row - wnd : tblRng.FirstRow;
                to = row + wnd < tblRng.LastRow ? row + wnd : tblRng.LastRow;
            }

            var ws = cell.Worksheet;
            Range col = ws.Range[ws.Cells[from, cell.Column], ws.Cells[to, cell.Column]];
            var tmp = new List<ExCell>();
            foreach (object r in col)
                tmp.Add(new ExCell((Range)r, false));
            CellValue[]  valList = tmp
                .Where(c => c.Value.Type == type &&
                            !string.IsNullOrEmpty(c.Value.ToString()))
                            .Select(c => c.Value).Distinct().ToArray();
            return valList;
        }
        public event PropertyChangedEventHandler PropertyChanged;

        [NotifyPropertyChangedInvocator]
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
