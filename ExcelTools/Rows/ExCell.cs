﻿using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Runtime.CompilerServices;
using ExcelTools.Annotations;
using Microsoft.Office.Interop.Excel;

namespace ExcelTools
{
    public class ExCell : INotifyPropertyChanged
    {
        private string _colName;
        private Range _rng;
        private bool _isSelected;
        public CellValue Value { get; }
        public CellValue[] ValList { get; private set; }
        public Range Rng { get { return _rng; } }

        public string ColName
        {
            get { return _colName; }
            set { _colName = value; }
        }
        public bool IsValid
        {
            get
            {
                if (_rng.Validation.Value) return true;
                return false;
            }
        }
        public bool IsSelected
        {
            get { return _isSelected; }
            set
            {
                if (value == IsSelected) return;
                _isSelected = value;
                OnPropertyChanged();
            }
        }
        public ExCell(Range rng, bool setValList)
        {
            _rng = rng;
            Value = new CellValue(_rng);
            Value.PropertyChanged += (sender, args) =>
            {
                if (args.PropertyName == "XlVal")
                    _rng.Value = Value.XlVal;
            };
            var cr = Current.CurRegion;
            Range colNameRng = (Range)cr.ActiveWs.Cells[cr.TblRange.FirstRow - 1, rng.Column];
            _colName = colNameRng.Value != null ? colNameRng.Value.ToString() : "";
            if (!setValList) return;
            int row = rng.Row;
            int wnd = ValListGenerator.Instance.Wnd;
            int from, to;
            if (ValListGenerator.Instance.AllRows)
            { 
                from = cr.TblRange.FirstRow;
                to = cr.TblRange.LastRow;
            }
            else
            {
                from = row - wnd > cr.TblRange.FirstRow ? row - wnd : cr.TblRange.FirstRow;
                to = row + wnd < cr.TblRange.LastRow ? row + wnd : cr.TblRange.LastRow;                
            }

            var ws = rng.Worksheet;
            Range col = ws.Range[ws.Cells[from, rng.Column], ws.Cells[to, rng.Column]];
            var tmp = new List<ExCell>();
            foreach (object r in col)
                tmp.Add(new ExCell((Range) r, false));
            ValList = tmp
                .Where(c=>c.Value.Type == Value.Type && 
                !string.IsNullOrEmpty(c.Value.ToString())).
                Select(c=>c.Value)
                .Distinct().ToArray();
        }

        public override string ToString()
        {
            return Value.StrVal;
        }

        public event PropertyChangedEventHandler PropertyChanged;

        [NotifyPropertyChangedInvocator]
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    public class CellValEquilityComparer : IEqualityComparer<ExCell>
    {
        public bool Equals(ExCell x, ExCell y)
        {
            return x.Value.ToString() == y.Value.ToString();
        }

        public int GetHashCode(ExCell obj)
        {
            return obj.Value.ToString().GetHashCode();
        }
    }
}