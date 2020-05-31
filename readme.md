# NPOI.Extensions
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

