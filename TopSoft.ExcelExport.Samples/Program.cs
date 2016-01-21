using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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
