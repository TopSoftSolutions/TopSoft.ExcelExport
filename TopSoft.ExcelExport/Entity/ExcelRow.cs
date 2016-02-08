using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using TopSoft.ExcelExport.Attributes;
using TopSoft.ExcelExport.Helpers;
using TopSoft.ExcelExport.Styles;

namespace TopSoft.ExcelExport.Entity
{
    public abstract class ExcelRow
    {
        private Dictionary<string, string> _propertyMappings = new Dictionary<string, string>();

        private Dictionary<string, List<object>> _propertyStylMappings = new Dictionary<string, List<object>>();

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

        public void MapStyle<T>(Expression<Func<T, object>> lambda, object style)
        {
            if(style == null)
            {
                throw new ArgumentNullException("style");
            }

            if(!(style is CellFill || style is CellText || style is CellBorder))
            {
                throw new ArgumentException("Unsupported Style Type.");
            }

            var member = lambda.Body as MemberExpression;
            if(member != null)
            {
                var propInfo = member.Member as PropertyInfo;
                if(propInfo != null)
                {
                    if(_propertyStylMappings.ContainsKey(propInfo.Name))
                    {
                        _propertyStylMappings[propInfo.Name].Add(style);
                    }
                    else
                    {
                        var newStylesList = new List<object>();
                        newStylesList.Add(style);
                        _propertyStylMappings.Add(propInfo.Name, newStylesList);
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
                if(target == null) { continue; }

                var dataCellAttr = dataCell.GetCustomAttributes(false).Where(atr => atr is CellDataAttribute).FirstOrDefault() as CellDataAttribute;
                if(dataCellAttr == null) { continue; }
 
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

                var excelCell = new ExcelCell();
                excelCell.Cell = cell;

                var cellFillAttr = dataCell.GetCustomAttributes(false).Where(atr => atr is CellFillAttribute).FirstOrDefault() as CellFillAttribute;
                var cellBorderAttr = dataCell.GetCustomAttributes(false).Where(atr => atr is CellBorderAttribute).FirstOrDefault() as CellBorderAttribute;
                var cellFontAttr = dataCell.GetCustomAttributes(false).Where(atr => atr is CellTextAttribute).FirstOrDefault() as CellTextAttribute;

                var cellStyleMappings = _propertyStylMappings.ContainsKey(dataCell.Name) ? _propertyStylMappings[dataCell.Name] : null;

                CellFill cellFillMap = null;
                CellBorder cellBorderMap = null;
                CellText cellFontMap = null;

                if(cellStyleMappings != null)
                {
                    cellFillMap = cellStyleMappings.Where(x => x is CellFill) as CellFill;
                    cellBorderMap = cellStyleMappings.Where(x => x is CellBorder) as CellBorder;
                    cellFontMap = cellStyleMappings.Where(x => x is CellText) as CellText;
                }

                
                if(cellFillMap == null)
                {
                    if(cellFillAttr != null) { excelCell.Styles.Add(cellFillAttr.GetFill()); }
                } else
                {
                    excelCell.Styles.Add(cellFillMap.GetFill());
                }

                if(cellBorderMap == null)
                {
                    if(cellBorderAttr != null) { excelCell.Styles.Add(cellBorderAttr.GetBorder()); }
                } else
                {
                    excelCell.Styles.Add(cellBorderMap.GetBorder());
                }

            
                if(cellFontMap == null)
                {
                    if(cellFontAttr != null) { excelCell.Styles.Add(cellFontAttr.GetFont()); }
                } else
                {
                    excelCell.Styles.Add(cellFontMap.GetFont());
                }


                retRowData.Cells.Add(excelCell);
            }

            return retRowData;
        }
    }
}
