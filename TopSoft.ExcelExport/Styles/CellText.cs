using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TopSoft.ExcelExport.Styles
{
    public class CellText
    {
        public bool Bold { get; set; }
        public bool Italic { get; set; }
        public bool Underliine { get; set; }

        public CellText(bool bold = false, bool italic = false, bool underline = false)
        {
            Bold = bold;
            Italic = italic;
            Underliine = underline;
        }
    }
}
