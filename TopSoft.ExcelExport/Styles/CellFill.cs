using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TopSoft.ExcelExport.Styles
{
    public class CellFill
    {
        public string HexColor { get; set; }

        public CellFill(string hexColor)
        {
            HexColor = hexColor;
        }
    }
}
