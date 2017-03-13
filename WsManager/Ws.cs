using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using WsManager.Annotations;
using Excel = Microsoft.Office.Interop.Excel;
namespace WsManager
{
    public class ExWs : INotifyPropertyChanged
    {
        public ExWs(Sheet ws, ExWb wb)
        {
            Ws = ws;
            Wb = wb;
        }
        public string Name { get { return Ws.Name; } }
        public ExWb Wb { get; }
        public Sheet Ws { get; }

        public WorksheetPart WsPart { get { return (WorksheetPart) Wb.SDoc.WorkbookPart.GetPartById(Ws.Id); } }


        private bool _isSelected;
        public bool IsSelected
        {
            get { return _isSelected; }
            set
            {
                if (_isSelected == value) return;
                _isSelected = value;
                OnPropertyChanged();
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        [NotifyPropertyChangedInvocator]
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
