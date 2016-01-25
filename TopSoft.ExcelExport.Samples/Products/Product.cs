using TopSoft.ExcelExport.Attributes;
using TopSoft.ExcelExport.Entity;

namespace TopSoft.ExcelExport.Samples.Products
{
    class Product : ExcelRow
    {
        [CellData("A"), CellBorder(left: true, right: true, top: true, bottom: true)]
        public string Name { get; set; }

        [CellData("B"), CellText(bold: true, italic: true)]
        public string Description { get; set; }

        [CellData("C"), CellFill(hexColor: "FFFF0000")]
        public double Price { get; set; }
    }
}
