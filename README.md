#### TopSoft.ExcelExport
~~Small~~ Very Small Toolkit to easy exporting data to excel

#### External Dependencies
> DocumentFormat.OpenXML

You can install it by running `Install-Package DocumentFormat.OpenXml` in Nu-Get Package Manager.

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

#### 2. Add `CellData` attribute to `Product` class properties to specify columns where they must be placed.
(Later we will show how to change them dynamically)

```c#
  class Product : ExcelRow
  {
      [DataCell("A")]
      public string Name { get; set; }

      [DataCell("B")]
      public string Description { get; set; }

      [DataCell("C")]
      public decimal Price { get; set; }
  }
```

#### 3. Create new excel or open existing one.

```c#
SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook);
```

#### 4 Initialize `ExcelExportContext` object with `SpreadsheetDocument.` 
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

That's All!

P.S. Toolkit is now in development stage. We will bee ready ~~soon~~ very soon.
