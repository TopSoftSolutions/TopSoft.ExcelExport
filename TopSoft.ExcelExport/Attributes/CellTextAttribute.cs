using System;

namespace TopSoft.ExcelExport.Attributes
{
    public class CellTextAttribute : Attribute
    {
        public bool Bold { get; private set; } 
        public bool Italic { get; private set; }
        public bool Underliine { get; private set; }

        public CellTextAttribute(bool bold = false, bool italic = false, bool underline = false)
        {
            Bold = bold;
            Italic = italic;
            Underliine = underline;
        }
    }
}
