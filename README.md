#### TopSoft.ExcelExport
~~Small~~ Very Small Toolkit to easy exporting data to excel

#### External Dependencies
> DocumentFormat.OpenXML

You can install it by running `Install-Package DocumentFormat.OpenXml` in Nu-Get Package Manager.

#### Restrictions
`Topsoft.ExcelExport` now works only with simple data types in models.

#### Let's Start

For example we have `Product` class and we want to export it to excel.

```c#
class Product
{
    public string Name { get; set; }
    public string Description { get; set; }
    public decimal Price { get; set; }
}
```
#### 1. Inherit `ExcelRow` in `Product` class.

```c#
class Product : ExcelRow
```

#### 2. Add `CellData` attribute to `Product` class properties to specify excel column names where they must be placed.
(Later we will show how to change them on the fly)

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

#### 3. Create new excel or open existing one.

```c#
SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook);
```

#### 4. Initialize `ExcelExportContext` object with `SpreadsheetDocument.` 
Just Call `RenderEntity` for each entity you want to appear in excel.

```c#
      var excelExportContext = new ExportContext(spreadsheetDocument)
      uint rowNo = 0;
      foreach(var product in products)
      {
          rowNo++;
          excelExportContext.RenderEntity(product, rowNo);
      }
```

#### 5. Adding excel column mappings on the fly.
You can add excel column mapping on the fly, wherever you want, before calling `RenderEntity` for model. Column Mappings are instance-level and will affect only object for which `MapColumn` was called.

For example:

```c#
   if(product.Price > 44)
   {
        product.MapColumn<Product>(x => x.Description, "F");
   }
```

In this example only this `product` 's Description will be placed at column "F". Other ones will be placed at their initial place ( specified by `CellData` attribute.


That's All!
