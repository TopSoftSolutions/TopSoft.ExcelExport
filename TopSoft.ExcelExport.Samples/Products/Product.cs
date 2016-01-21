using TopSoft.ExcelExport.Attributes;
using TopSoft.ExcelExport.Entity;

namespace TopSoft.ExcelExport.Samples.Products
{
    class Product : ExcelRow
    {
        [CellData("A")]
        public string Name { get; set; }

        [CellData("B")]
        public string Description { get; set; }

        [CellData("C")]
        public double Price { get; set; }
    }
}
