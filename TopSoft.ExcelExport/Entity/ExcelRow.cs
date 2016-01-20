using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using TopSoft.ExcelExport.Attributes;
using TopSoft.ExcelExport.Helpers;

namespace TopSoft.ExcelExport.Entity
{
    public abstract class ExcelRow
    {
        public RowData ToRow(uint rowNo)
        {
            if(rowNo == 0)
            {
                throw new ArgumentException("rowNo must be greater than 0.");
            }

            var retRowData = new RowData();

            retRowData.Row.RowIndex = rowNo;

            var dataRowType = GetType();
            var dataCellProperties = dataRowType.GetProperties().Where(prop => prop.IsDefined(typeof(CellDataAttribute), false));

            foreach(var dataCell in dataCellProperties)
            {
                if(dataCell == null) { continue; }

                var target = dataCell.GetValue(this);
                if(target != null)
                {
                    var dataCellAttr = dataCell.GetCustomAttributes(false).Where(atr => atr is CellDataAttribute).FirstOrDefault() as CellDataAttribute;
                    if(dataCellAttr != null)
                    {
                        var cellColumnName = dataCellAttr.ColumnName;

                        var cellDataType = ExcelHelper.ResolveCellType(target.GetType());
                        var cellDataValue = target.ToString();

                        var cellReference = cellColumnName + rowNo;
                        var cell = new Cell() { CellReference = cellReference };

                        cell.CellValue = new CellValue(cellDataValue);
                        cell.DataType = new EnumValue<CellValues>(cellDataType);

                        retRowData.Cells.Add(cell);
                    }
                }
            }

            return retRowData;
        }
    }
}
