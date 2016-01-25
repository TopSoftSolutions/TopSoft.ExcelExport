using TopSoft.ExcelExport.Samples.Products;

namespace TopSoft.ExcelExport.Samples
{
    class Program
    {
        static void Main()
        {
            var productExport = new ProductExport();
            productExport.Process();
        }
    }
}
