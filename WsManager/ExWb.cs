using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using WsManager.Annotations;
using Excel = Microsoft.Office.Interop.Excel;
using Path = System.IO.Path;

namespace WsManager
{
    public class ExWb 
    {
        private SpreadsheetDocument _sDoc;
        public List<Sheet> Sheets { get; }
        public ExWb(string file)
        {
            File = file;
            WorkSheets = new List<ExWs>();
            Sheets = new List<Sheet>();
            Open(true);
            ReadWSheets();
            Close();
        }
        public SpreadsheetDocument SDoc
        {
            get { return _sDoc; }
            set { _sDoc = value; }
        }

        public string File { get; set; }
        public string FileName { get { return Path.GetFileName(File); } }
        public string DirName { get { return Path.GetDirectoryName(File); } }
        public bool IsOpened { get { return _sDoc == null; } }
        public List<ExWs> WorkSheets { get;  }

        public string GetSharedStr(string idStr)
        {
            int id;
            if (!int.TryParse(idStr, out id)) return "";
            var items = _sDoc.WorkbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ToArray();
            if (id >= items.Length) return "";
            return items[id].Text.Text;

        }

        public void Open(bool write = false)
        {
            SDoc = SpreadsheetDocument.Open(File, write);
            ReadWSheets();
        }

        public void Close()
        {
            SDoc.Close();
        }
        static WorksheetPart GetWorkSheetPart(WorkbookPart workbookPart, string sheetName)
        {
            string relId = workbookPart.Workbook
            .Descendants<Sheet>()
            .First(s => s.Name.Value.Equals(sheetName)).Id;
            return (WorksheetPart)workbookPart.GetPartById(relId);
        }
        static void CleanView(WorksheetPart worksheetPart)
        {
            //There can only be one sheet that has focus
            SheetViews views = worksheetPart.Worksheet.GetFirstChild<SheetViews>();
            if (views != null)
            {
                views.Remove();
                worksheetPart.Worksheet.Save();
            }
        }

        public void Add(ExWs ws, string name)
        {
            try
            {
                
                SpreadsheetDocument tempDoc = SpreadsheetDocument.Create(new MemoryStream(), SDoc.DocumentType);
                WorkbookPart tempWorkbookPart = tempDoc.AddWorkbookPart();
                //SharedStringTablePart sharedStringTable = ws.Wb.SDoc.WorkbookPart.SharedStringTablePart;
                //SharedStringTablePart tempSharedStringTable = tempWorkbookPart.AddPart(sharedStringTable);
                WorksheetPart tempWorksheetPart = tempWorkbookPart.AddPart(ws.WsPart);                       
                WorksheetPart clonedSheet = SDoc.WorkbookPart.AddPart(tempWorksheetPart);
                var sData = clonedSheet.Worksheet.Elements<SheetData>().First();
                var sourceNumberingFormats = ws.Wb.SDoc.WorkbookPart.WorkbookStylesPart.Stylesheet.NumberingFormats;
                foreach (var row in sData.Elements<Row>())
                {
                    foreach (var cell in row.Elements<Cell>())
                    {
                        if (cell.DataType != null && cell.DataType == CellValues.SharedString)
                        {
                            string str = ws.Wb.GetSharedStr(cell.CellValue.Text);
                            cell.DataType = CellValues.String;
                            cell.CellValue = new CellValue(str);
                        }
                            
                    }
                }
                //clonedSheet.AddPart(tempSharedStringTable);
                //Table definition parts are somewhat special and need unique ids…so let’s make an id based on count
                int numTableDefParts = ws.WsPart.GetPartsCountOfType<TableDefinitionPart>();
                int tableId = numTableDefParts;
                //Clean up table definition parts (tables need unique ids)
                if (numTableDefParts != 0)
                {
                    //Every table needs a unique id and name
                    foreach (TableDefinitionPart tableDefPart in clonedSheet.TableDefinitionParts)
                    {
                        tableId++;
                        tableDefPart.Table.Id = (uint)tableId;
                        tableDefPart.Table.DisplayName = "CopiedTable" +tableId;
                        tableDefPart.Table.Name = "CopiedTable" +tableId;
                        tableDefPart.Table.Save();
                    }
                }

                //There should only be one sheet that has focus

                CleanView(clonedSheet);
                var wbPart = SDoc.WorkbookPart;
                Sheets sheets = wbPart.Workbook.GetFirstChild<Sheets>();
                Sheet copiedSheet = new Sheet();
                copiedSheet.Name = name;
                copiedSheet.Id = wbPart.GetIdOfPart(clonedSheet);
                copiedSheet.SheetId = (uint)sheets.ChildElements.Count + 1;
                sheets.Append(copiedSheet);
                wbPart.Workbook.Save();
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                //throw;
            }

        }

        private void ReadWSheets()
        {
            WorkSheets.Clear();
            foreach (Sheet sheet in SDoc.WorkbookPart.Workbook.Descendants<Sheet>())
            {
                Sheets.Add( sheet);
                WorkSheets.Add(new ExWs(sheet, this));
            }
        }


    }
}
