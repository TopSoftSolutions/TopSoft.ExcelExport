using System;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using TopSoft.ExcelExport.Entity;
using TopSoft.ExcelExport.Helpers;

namespace TopSoft.ExcelExport
{
    public class ExportContext
    {
        public SpreadsheetDocument SpreadSheet { get; private set; }
        public Worksheet Worksheet { get; private set; }
        public SheetData SheetData { get; private set; }

        public ExportContext(SpreadsheetDocument spreadSheet)
        {
            if(spreadSheet == null)
            {
                throw new ArgumentNullException("spreadSheet");
            }

            // TODO check spreadSHeet if opened with canEdit flag
            // TODO check worksheet, if not exist create

            SpreadSheet = spreadSheet;
            Worksheet = SpreadSheet.WorkbookPart.WorksheetParts.First().Worksheet;
            SheetData = Worksheet.GetFirstChild<SheetData>();

        }

        public void RenderEntity(object entity, uint rowNo)
        {
            var excelRow = entity as ExcelRow;
            if(excelRow != null)
            {
                RowData rowData = excelRow.ToRow(rowNo);

                foreach(var cell in rowData.Cells)
                {
                    var existingCell = rowData.Row.Elements<Cell>().Where(e => e.CellReference == cell.CellReference).FirstOrDefault();

                    if(existingCell == null)
                    {
                        var nextCell = rowData.Row.Elements<Cell>().Where(e => ExcelHelper.ColumnCompare(e.CellReference.Value, cell.CellReference) > 0).FirstOrDefault();
                        rowData.Row.InsertBefore(cell, nextCell);
                    }
                    else
                    {
                        existingCell.CellValue = cell.CellValue;
                        existingCell.DataType = cell.DataType;
                    }
                }

                SheetData.Append(rowData.Row);
            }
        }
    }
}
