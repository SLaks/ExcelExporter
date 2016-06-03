#ExcelExport

ExcelExport is a simple, fluent API to export data to Excel spreadsheets.  ExcelExport uses OleDb to generate Excel files (Microsoft.ACE.OLEDB for Excel 2007+ formats, and Microsoft.Jet.OLEDB for Excel 2003 `.xls` files)

This library can also be used to generate Excel files in ASP.Net MVC actions; use [this simple ActionResult class](https://gist.github.com/3044898).

##Sample usage

```C#
new ExcelExport()
	.AddSheet("Sample Names", new[] {
		new { Name = "Bill Stewart",	ZipCode = "00347", Birth_Date = new DateTime(1987, 6, 5) },
		new { Name = "Russ Porter",  	ZipCode = "04257", Birth_Date = new DateTime(1956, 7, 8) },
		new { Name = "Rodrick Rivers",	ZipCode = "19867", Birth_Date = new DateTime(1956, 7, 8) }
	})
	.AddSheet(
		"LINQ Query Sample",
		ordersQuery.Select(o => new { 
			Product_Name = o.Product.Name, 
			o.Quantity,
			o.OrderDate
		}
	)
	.AddSheet(someDataSet.Tables[0])
	.AddSheet(
		"Classic ADO.Net Sample",
		someCommand.ExecuteReader()
	)
	.ExportTo(Path.Combine(
		Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), 
		"Sample.xlsx"
	));
```

##Notes

 - When exporting anonymous types, `_` (underscore) characters in property names will be replaced with spaces.

 - When exporting ADO.Net DataTables, the sheet name is optional; if omitted, the table's `TableName` property will be used instead.

 - When exporting ADO.Net DataReaders, the reader must remain open when `ExportTo()` is called.  When the export is finished, the reader will be closed.
 - 
 

##License

ExcelExport is Copyright © 2016 by Contributors under the MIT license (https://opensource.org/licenses/MIT).
