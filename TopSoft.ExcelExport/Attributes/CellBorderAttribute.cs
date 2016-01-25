using System;

namespace TopSoft.ExcelExport.Attributes
{
    public class CellBorderAttribute : Attribute
    {
        public bool LeftBorder { get; private set; }
        public bool RightBorder  { get; private set; }
        public bool TopBorder  { get; private set; }
        public bool BottomBorder  { get; private set; }
        public bool DiagonalBorder  { get; private set; }

        public CellBorderAttribute(bool left = false,
            bool right = false, 
            bool top = false,
            bool bottom = false, 
            bool diagonal = false)
        {
            LeftBorder = left;
            RightBorder = right;
            TopBorder = top;
            BottomBorder = bottom;
            DiagonalBorder = diagonal;
        }
    }
}
