using System;

namespace TopSoft.ExcelExport.Attributes
{
    class CellDataAttribute : Attribute
    {
        public string ColumnName { get; private set; }
        public CellDataAttribute(string columnName)
        {
            ColumnName = columnName;
        }
    }
}
