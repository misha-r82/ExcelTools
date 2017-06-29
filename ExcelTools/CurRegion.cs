using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using ExcelTools.Annotations;
using Microsoft.Office.Interop.Excel;

namespace ExcelTools
{
    public static class Current
    {
        public static CurRegion CurRegion { get; }
        static Current()
        {
            CurRegion = new CurRegion();
            CurRegion.Reload();
            CurRegion.Wnd = 50;
        }
    }
    public class CurRegion : INotifyPropertyChanged
    {
        public CurRegion()
        {
            ActiveWs.Application.SheetSelectionChange += ApplicationOnSheetSelectionChange;
        }

        public int firstRow, lastRow, firstCol, lastCol;
        private Range _selection;
        private Range _activeCell;
        private Range _curRng;
        private ActiveRow _activeRow;
        private int _wnd;
        private bool _isWorkSheet;
        private bool _isTableCell;


        public Worksheet ActiveWs { get { return (Worksheet)ThisWorkbook.app.ActiveSheet; } }
        public Range Selection { get { return _selection; } }
        public Range ActiveCell { get { return _activeCell; } }

        public bool IsWorkSheet
        {
            get { return _isWorkSheet; }
            private set
            {
                if (_isWorkSheet == value) return;
                _isWorkSheet = value;
                OnPropertyChanged();
            }
        }

        public ActiveRow ActiveRow
        {
            get
            {
                /*if (_activeRow == null) Debug.WriteLine("null");
                else Debug.WriteLine(_activeRow.Cells[0]);*/
                return _activeRow;
            }
        }

        public Range CurRng { get { return _curRng; } }
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
        public bool IsTableCell
        {
            get { return _isTableCell; }
            set
            {
                if (value == _isTableCell) return;
                _isTableCell = value;
                OnPropertyChanged();
            }
        }

        public int CurRowNumInRng
        {
            get
            {
                if (_activeRow == null) return -1;
                return _activeRow.RowRng.Row - firstRow + 1;
            }
            set
            {
                if (value > 0 && value <= lastRow)
                {
                    var cell = (Range)ActiveWs.Cells[value + firstRow -1, ActiveCell.Column];
                    cell.Select();
                }
                
                OnPropertyChanged();
            }
        }

        public void Reload()
        {
            /*try
            {
                var tmp = (PivotTable)ActiveWs.PivotTables(1);
                var fields = (PivotFields)tmp.PivotFields();
                for (int i = 1; i <= fields.Count; i++)
                {
                    var field = (PivotField)fields.Item(i);
                    XlPivotFieldDataType type = field.DataType;
                    Debug.WriteLine(field.Name + "-"+type);
                }
            }
            catch (Exception e)
            {

            }*/
            
            _selection = (Range)ThisWorkbook.app.Selection;
            _activeCell = ThisWorkbook.app.ActiveCell;
            _curRng = _activeCell.CurrentRegion;
            Range firstCell = (Range)_curRng[2, 1];
            firstRow = firstCell.Row;
            firstCol = firstCell.Column;
            lastRow = firstRow + _curRng.Rows.Count -2;
            lastCol = firstCol + _curRng.Columns.Count;
            _activeRow = IsTableCell ? new ActiveRow(ActiveCell) : null;
            OnPropertyChanged(nameof(ExcelTools.ActiveRow));
            OnPropertyChanged(nameof(CurRowNumInRng));               
            
        }
        private void ApplicationOnSheetSelectionChange(object sh, Range target)
        {
            var sheet =  sh as Worksheet;
            IsWorkSheet = sheet.Type == XlSheetType.xlWorksheet;
            IsTableCell = ((PivotTables)sheet.PivotTables()).Count == 0;
            if (!IsWorkSheet) return;
            Reload();
        }

        public event PropertyChangedEventHandler PropertyChanged;

        [NotifyPropertyChangedInvocator]
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
