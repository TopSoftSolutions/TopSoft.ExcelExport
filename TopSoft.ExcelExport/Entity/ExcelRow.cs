using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using TopSoft.ExcelExport.Attributes;
using TopSoft.ExcelExport.Helpers;

namespace TopSoft.ExcelExport.Entity
{
    public abstract class ExcelRow
    {
        private Dictionary<string, string> _propertyMappings = new Dictionary<string, string>();

        public void MapColumn<T>(Expression<Func<T, object>> lambda, string columnName)
        {
            if(string.IsNullOrEmpty(columnName))
            {
                throw new ArgumentNullException(columnName);
            }

            var member = lambda.Body as MemberExpression;
            if(member != null)
            {
                var propInfo = member.Member as PropertyInfo;
                if(propInfo != null)
                {
                    if(_propertyMappings.ContainsKey(propInfo.Name))
                    {
                        _propertyMappings[propInfo.Name] = columnName;
                    }
                    else
                    {
                        _propertyMappings.Add(propInfo.Name, columnName);
                    }
                }
            }
        }

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

                        if(_propertyMappings.ContainsKey(dataCell.Name)) 
                        {
                            cellColumnName = _propertyMappings[dataCell.Name];
                        }

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
