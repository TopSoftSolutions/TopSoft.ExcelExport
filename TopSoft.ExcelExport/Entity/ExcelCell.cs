using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace TopSoft.ExcelExport.Entity
{
    public class ExcelCell
    {
        public List<OpenXmlElement> Styles { get; set; }
        public Cell Cell { get; set; }

        public ExcelCell()
        {
            Styles = new List<OpenXmlElement>();
            Cell = new Cell();
        }
    }
}
