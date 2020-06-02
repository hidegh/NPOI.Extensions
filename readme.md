# NPOI.Extensions

![Nuget](https://img.shields.io/nuget/v/NPOI.Extensions?label=NuGet%3A%20NPOI.Extensions%20%28make%20sure%20you%20use%20plural%29&logo=nuget)

This package is an extension to the https://github.com/tonyqus/npoi, as I found reading excel values with the basic interface a bit cumbersome.

> On my pet projects I used less flexible helper functions for reading excel cell values. They always adressed just the needs of a given use-case and were not tested for flexibility and corner-case cenarios.

**My goal with this library is to have a stable, reliable, simple (unified) way to read values from excel cell.**
To ensure this, project includes unit tests for the GetValue extension method, covering most of the situations a developer might need in an enterprise solution.


## This library currently

- **simplifies & unifies reading cell values into typed variables**
- adds support to **map a sheet into a typed list** - *simple approach with 100% control over the mapping*

*NOTE:*

There's also a project (https://github.com/donnytian/Npoi.Mapper) which addresses mapping sheet to typed list, which requires a major change in the Map and ShouldMap function to be more widely usable (see: https://github.com/donnytian/Npoi.Mapper/issues/45).

## Major extension functions

**public T GetValue<T>(this IRow, index)**
- **unifies access to cell value regardless of it's type**
- allows to read into a typed output variable in a simple way
- does a **safe null/blank cell access**, is able to map cell value (being it null or not) to Nullable&lt;&gt; types
- allows to **read any cell as a formatted string** (the formatted content - that is shown inside the excel cell - is returned as a string)
- handles reading of **formula results** into typed variable (we usually are interested in the result, not in the formula)

---
**public List<T> MapTo<T>(this ISheet sheet, bool firstRowIsHeader, Func<ExcelRowMapper, T> rowMapper)**
- **simplifies mapping the excel content into a typed list**
- the rowMapperFunction allows to **get cell content by column header (title)**: rowMapper.GetValue&lt;DateTime&gt;("Order Date")


## Getting started

There's a **sample project** (and sample excel workbooks) for the mappign extensions, where also the GetValue<T> is used.

The **XUnitTest** for the GetValueExtensions contains tests for all the functionality the GetValue<T> is intended to support.

## License
https://opensource.org/licenses/MIT, Copyright (c) 2020 by Balázs HIDEGHÉTY

## Sample

```csharp
public class Sample
{
	protected const string testExcelFileName = "./MapExtensions/Sample.xlsx";

	public class SheetColumnTitles
	{
		public const string Date = "Order Date";
		public const string Region = "Region";
		public const string Representative = "Representative";
		public const string Item = "Item";
		public const string Units = "Units";
		public const string UnitCost = "Unit Cost";
		public const string Total = "Total";
	}

	public enum RegionEnum
	{
		East,
		Central,
		West
	}

	public class OrderDetails
	{
		public DateTime Date { get; set; }
		public RegionEnum Region { get; set; }
		public string Representative { get; set; }
		public string Item { get; set; }
		public int Units { get; set; }
		public decimal UnitCost { get; set; }
		public decimal Total { get; set; }

		public string DateFormatted { get; set; }
		public string TotalFormatted { get; set; }
	}

	public void Run_MapToClass()
	{
		IWorkbook workbook;

		using (var fileReader = new FileStream(testExcelFileName, FileMode.Open, FileAccess.Read))
			workbook = new XSSFWorkbook(fileReader);

		var sheet = workbook.GetSheetAt(0);

		// reusable mapping function from value to enum...
		var regionMapper = (Func<string, RegionEnum>)((string region) =>
		{
			switch ((region ?? "").ToUpper())
			{
				case "EAST": return RegionEnum.East;
				case "CENTRAL": return RegionEnum.Central;
				case "WEST": return RegionEnum.West;
				default: throw new NotSupportedException();
			}
		});

		// map to list
		var data = sheet.MapTo<OrderDetails>(true, rowMapper =>
		{
			// map singleItem
			return new OrderDetails()
			{
				Date = rowMapper.GetValue<DateTime>(SheetColumnTitles.Date),
        
				// use reusable mapper for re-curring scenarios
				Region = regionMapper(rowMapper.GetValue<string>(SheetColumnTitles.Region)),
        
				Representative = rowMapper.GetValue<string>(SheetColumnTitles.Representative),
				Item = rowMapper.GetValue<string>(SheetColumnTitles.Item),
				Units = rowMapper.GetValue<int>(SheetColumnTitles.Units),
				UnitCost = rowMapper.GetValue<decimal>(SheetColumnTitles.UnitCost),
				Total = rowMapper.GetValue<decimal>(SheetColumnTitles.Total),
        
				// read date and total as string, as they're displayed/formatted on the excel
				DateFormatted = rowMapper.GetValue<string>(SheetColumnTitles.Date),
				TotalFormatted = rowMapper.GetValue<string>(SheetColumnTitles.Total)
			};
		});

		// write to console
		Console.WriteLine("Date       | Region  | Representative  | Item            | Units | Unit cost  | Total      |");
		Console.WriteLine("-------------------------------------------------------------------------------------------- ");

		data.ForEach(data =>
		{
			Console.WriteLine("{0:yyyy-MM-dd} | {1,-7} | {2,-15} | {3,-15} | {4,5:N0} | {5,10:N2} | {6,10:N2} | {7,10}, {8}", data.Date, data.Region, data.Representative, data.Item, data.Units, data.UnitCost, data.Total, data.DateFormatted, data.TotalFormatted);
		});

	}

}
```
