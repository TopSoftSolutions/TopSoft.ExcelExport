namespace TopSoft.ExcelExport.Styles
{
    public class CellBorder
    {
        public bool LeftBorder { get; set; }
        public bool RightBorder { get; set; }
        public bool TopBorder { get; set; }
        public bool BottomBorder { get; set; }
        public bool DiagonalBorder { get; set; }

        public CellBorder(bool left = false,
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
