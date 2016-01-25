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
        public Stylesheet StyleSheet { get; private set; }

        public ExportContext(SpreadsheetDocument spreadSheet)
        {
            if(spreadSheet == null)
            {
                throw new ArgumentNullException("spreadSheet");
            }

            if(spreadSheet.FileOpenAccess != System.IO.FileAccess.ReadWrite)
            {
                throw new Exception("No access granted for opened excel");
            }

            SpreadSheet = spreadSheet;

            if(SpreadSheet.WorkbookPart == null)
            {
                // Add a WorkbookPart to the document.
                WorkbookPart workbookpart = SpreadSheet.AddWorkbookPart();
                workbookpart.Workbook = new Workbook();

                // Add a WorksheetPart to the WorkbookPart.
                WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());

                // Add Sheets to the Workbook.
                Sheets sheets = SpreadSheet.WorkbookPart.Workbook.
                    AppendChild(new Sheets());

                // Append a new worksheet and associate it with the workbook.
                Sheet sheet = new Sheet()
                {
                    Id = SpreadSheet.WorkbookPart.
                    GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = "Sheet 1"
                };
                sheets.Append(sheet);

                workbookpart.Workbook.Save();

            }

            WorkbookStylesPart stylesPart = SpreadSheet.WorkbookPart.AddNewPart<WorkbookStylesPart>();
            stylesPart.Stylesheet = new Stylesheet();
            StyleSheet = stylesPart.Stylesheet;

            StyleSheet.Fonts = new Fonts();
            // required by Excel
            StyleSheet.Fonts.AppendChild(new Font());
            StyleSheet.Fonts.Count = 1;

            StyleSheet.Fills = new Fills();
            // required, reserved by Excel
            StyleSheet.Fills.Append(new Fill { PatternFill = new PatternFill { PatternType = PatternValues.None } });
            // required, reserved by Excel 
            StyleSheet.Fills.Append(new Fill { PatternFill = new PatternFill { PatternType = PatternValues.Gray125 } }); 
            StyleSheet.Fills.Count = 2;

            StyleSheet.Borders = new Borders();
            // required by Excel
            StyleSheet.Borders.Append(new Border());
            StyleSheet.Borders.Count = 1;

            StyleSheet.CellStyleFormats = new CellStyleFormats();
            // required by Excel
            StyleSheet.CellStyleFormats.Append(new CellFormat());
            StyleSheet.CellStyleFormats.Count = 1;

            StyleSheet.CellFormats = new CellFormats();
            // required by Excel
            StyleSheet.CellFormats.AppendChild(new CellFormat());
            StyleSheet.CellFormats.Count = 1;

            Worksheet = SpreadSheet.WorkbookPart.WorksheetParts.First().Worksheet;
            SheetData = Worksheet.GetFirstChild<SheetData>();

        }

        public void RenderEntity(object entity, uint rowNo)
        {
            var excelRow = entity as ExcelRow;
            if(excelRow != null)
            {
                RowData rowData = excelRow.ToRow(rowNo);

                foreach(var exelCell in rowData.Cells)
                {
                    if(exelCell.Cell == null) { continue; }

                    var existingCell = rowData.Row.Elements<Cell>().Where(e => e.CellReference == exelCell.Cell.CellReference).FirstOrDefault();

                    if(existingCell == null)
                    {
                        var nextCell = rowData.Row.Elements<Cell>().Where(e => ExcelHelper.ColumnCompare(e.CellReference.Value, exelCell.Cell.CellReference) > 0).FirstOrDefault();
                        rowData.Row.InsertBefore(exelCell.Cell, nextCell);
                        existingCell = exelCell.Cell;
                    }
                    else
                    {
                        existingCell.CellValue = exelCell.Cell.CellValue;
                        existingCell.DataType = exelCell.Cell.DataType;
                    }

                    if(exelCell.Styles.Any())
                    {
                        CellFormat cellFormat = existingCell.StyleIndex != null ? GetCellFormat(existingCell.StyleIndex).CloneNode(true) as CellFormat : new CellFormat();

                        var fillStyle = exelCell.Styles.Where(x => x is PatternFill).FirstOrDefault() as PatternFill;
                        if(fillStyle != null)
                        {
                            cellFormat.FillId = InsertFill(fillStyle);
                            cellFormat.ApplyFill = true;
                        }

                        var borderStyle = exelCell.Styles.Where(x => x is Border).FirstOrDefault() as Border;
                        if(borderStyle != null)
                        {
                            cellFormat.BorderId = InsertBorder(borderStyle);
                            cellFormat.ApplyBorder = true;
                        }

                        var textStyle = exelCell.Styles.Where(x => x is Font).FirstOrDefault() as Font;
                        if(textStyle != null)
                        {
                            cellFormat.FontId = InsertFont(textStyle);
                            cellFormat.ApplyFont = true;
                        }

                        existingCell.StyleIndex = InsertCellFormat(cellFormat);
                    }

                }

                SheetData.Append(rowData.Row);
            }
        }

        public void SaveChanges()
        {
            StyleSheet.Save();
            Worksheet.Save();
        }

        private uint InsertFill(PatternFill fill)
        {
            StyleSheet.Fills.Append(new Fill { PatternFill = fill });
            StyleSheet.Fills.Count++;
            uint insertedIndex = StyleSheet.Fills.Count - 1;

            return insertedIndex;
        }

        private uint InsertFont(Font font)
        {
            StyleSheet.Fonts.Append(font);
            StyleSheet.Fonts.Count++;
            uint insertedIndex = StyleSheet.Fonts.Count - 1;

            return insertedIndex;
        }

        private uint InsertBorder(Border border)
        {
            StyleSheet.Borders.Append(border);
            StyleSheet.Borders.Count++;
            uint insertedIndex = StyleSheet.Borders.Count - 1;

            return insertedIndex;
        }

        private uint InsertCellFormat(CellFormat cellFormat)
        {
            StyleSheet.CellFormats.Append(cellFormat);
            StyleSheet.CellFormats.Count++;
            uint insertedIndex = StyleSheet.CellFormats.Count - 1;

            return insertedIndex;
        }

        private CellFormat GetCellFormat(uint styleIndex)
        {
            return StyleSheet.CellFormats.Elements<CellFormat>().ElementAt((int)styleIndex);
        }

    }
}
