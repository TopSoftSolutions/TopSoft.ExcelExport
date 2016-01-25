#### TopSoft.ExcelExport
Toolkit to easy exporting data to excel via C#

#### Install TopSoft.ExcelExport via NuGet
To install TopSoft.ExcelExport, run the following command in the Package Manager Console

`PM> Install-Package TopSoft.ExcelExport`

#### External Dependencies
> OpenXML SDK 2.5

You can install it by running `Install-Package DocumentFormat.OpenXml` in the Nu-Get Package Manager.

#### Restrictions
`Topsoft.ExcelExport` now works only for simple data types in models.
Styles can't be applied on the fly. ( in development )

#### Let's Start

Let's suppose we have the `Product` class that we want to export to Microsoft Excel document.

```c#
class Product
{
    public string Name { get; set; }
    public string Description { get; set; }
    public decimal Price { get; set; }
}
```
#### Step 1. Inherit `Product` class from the `ExcelRow` base class.

```c#
class Product : ExcelRow
```

#### Step 2. Add `CellData` attribute to the`Product` class's properties in-order to specify the document's column where the data should be placed.
(Later we will show how to change the column names on the fly.)

```c#
  class Product : ExcelRow
  {
      [CellData("A")]
      public string Name { get; set; }

      [CellData("B")]
      public string Description { get; set; }

      [CellData("C")]
      public decimal Price { get; set; }
  }
```

#### Step 3. Create or open an Excel document using the OpenXML SDK.

```c#
SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook);
```

#### Step 4. Initialize `ExcelExportContext` object by passing the instance of the target `SpreadsheetDocument.` Then by calling the `RenderEntity`  for each of the entities, they'll be exported to the Excel file.

```c#
      var excelExportContext = new ExportContext(spreadsheetDocument)
      uint rowNo = 0;
      foreach(var product in products)
      {
          rowNo++;
          excelExportContext.RenderEntity(product, rowNo);
      }
      excelExportContext.SaveChanges();
```

#### Step 5. Adding excel column mappings on the fly.
You can add excel column mapping on the fly before calling `RenderEntity` function for model entities. Column Mappings are instance-level and will affect only the particular object for which `MapColumn` function has been called.

For example:

```c#
   if(product.Price > 44)
   {
        product.MapColumn<Product>(x => x.Description, "F");
   }
```

In this example, the description will be placed at column "F" only for this particular instance of `product`. Other entities will continue to use column name specified by the `CellData` attribute.

#### Step 5. What about styles ?

Here's example of usage `CellBorder`, `CellText` and `CellFill` attributes:

```c#
    class Product : ExcelRow
    {
        [CellData("A"), CellBorder(left: true, right: true, top: true, bottom: true)]
        public string Name { get; set; }

        [CellData("B"), CellText(bold: true, italic: true)]
        public string Description { get; set; }

        [CellData("C"), CellFill(hexColor: "FFFF0000")]
        public double Price { get; set; }
    }
```


That's All!

View [Samples Project](https://github.com/TopSoftSolutions/TopSoft.ExcelExport/tree/master/TopSoft.ExcelExport.Samples) for more examples.

Read [wiki page](https://github.com/TopSoftSolutions/TopSoft.ExcelExport/wiki) for more info about exporting data to excel.

