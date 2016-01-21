using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace TopSoft.ExcelExport.Samples.Products
{
    class ProductExport
    {
        private List<Product> _products = new List<Product>()
        {
            new Product() { Name="telephone", Description="telephone sample description", Price = 10.5},
            new Product() { Name="tv", Description="tv sample description", Price = 22.5},
            new Product() { Name="notebook", Description="notebook sample description", Price = 44.66},
            new Product() { Name="monitor", Description="monitor sample description", Price = 77.8},
            new Product() { Name="keyboard", Description="keyboard sample description", Price = 90.5}

        };
        public void Process()
        {
            using(SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(@"D:\product.xlsx", SpreadsheetDocumentType.Workbook))
            {
                var excelExportContext = new ExportContext(spreadsheetDocument);

                uint rowNo = 0;
                foreach(var product in _products)
                {
                    rowNo++;
                    if(product.Price > 44)
                    {
                        product.MapColumn<Product>(x => x.Description, "F");
                    }

                    excelExportContext.RenderEntity(product, rowNo);
                }

                excelExportContext.SpreadSheet.WorkbookPart.Workbook.Save();
            }
        
        }
    }
}
