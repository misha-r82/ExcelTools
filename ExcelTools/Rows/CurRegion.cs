﻿using System;
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
            CurRegion.Init();            
        }
    }
    public class CurRegion : INotifyPropertyChanged
    {
        public CurRegion()
        {
            ActiveWs.Application.SheetSelectionChange += ApplicationOnSheetSelectionChange;
            ActiveWs.Application.SheetActivate += Application_SheetActivate;            
        }

        private Range _selection;
        private Range _activeCell;
        private ActiveRow _activeRow;
        private bool _isWorkSheet;
        private bool _isTableCell;
        private TblRange _tblRange;


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
                else Debug.WriteLine(_activeRow.ExCells[0]);*/
                return _activeRow;
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
                if (!IsTableCell) return -1;
                if (_activeRow == null) return -1;
                return _activeRow.RowRng.Row - TblRange.FirstRow + 1;
            }
            set
            {
                if (value > 0 && value <= TblRange.LastRow)
                {
                    var cell = (Range)ActiveWs.Cells[value + TblRange.FirstRow -1, ActiveCell.Column];
                    cell.Select();
                }
                
                OnPropertyChanged();
            }
        }

        public TblRange TblRange
        {
            get { return _tblRange; }
        }

        public void Init()
        {
            Application_SheetActivate(ActiveWs);
        }
        public void Reload()
        {
            if (!IsWorkSheet) return;
            _selection = (Range)ThisWorkbook.app.Selection;
            _activeCell = ThisWorkbook.app.ActiveCell;
            _tblRange = new TblRange(_activeCell);
            _activeRow = new ActiveRow(ActiveCell);
            OnPropertyChanged(nameof(ExcelTools.ActiveRow));
            OnPropertyChanged(nameof(CurRowNumInRng));               
            
        }
        private void Application_SheetActivate(object sh)
        {
            var sheet = sh as Worksheet;
            IsWorkSheet = sheet.Type == XlSheetType.xlWorksheet;
            IsTableCell = ((PivotTables)sheet.PivotTables()).Count == 0;
            Reload();
        }
        private void ApplicationOnSheetSelectionChange(object sh, Range target)
        {
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
