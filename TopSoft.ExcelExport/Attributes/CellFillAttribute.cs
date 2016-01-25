using System;

namespace TopSoft.ExcelExport.Attributes
{
    public class CellFillAttribute : Attribute
    {
        public string HexColor { get; private set; }

        public CellFillAttribute(string hexColor)
        {
            HexColor = hexColor;
        }
    }
}
