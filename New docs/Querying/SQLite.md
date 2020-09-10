# SQL in Excel
QueryStorm comes equipped with a [SQLite](https://www.sqlite.org) engine that can work with Excel tables as if they were database tables.

## Connecting
Clicking the SQL button in the ribbon pops up the QueryStorm IDE and creates a new SQLite script that you can use to run queries against the tables in the current workbook.

![Connect to workbook](../../Images/connect_sql.png)

## Querying

Once connected, we can start querying. It's important to note that the SQLite engine will only see data that is inside [Excel tables](http://www.excel-easy.com/data-analysis/tables.html "Excel tables intro"). 

> To turn a range into a table, select it and press `Ctrl+T`

![Querying](../../Images/sql_querying.gif?v=1)

You can use SQL to query and modify data inside tables. All four SQL data operations are supported: `select`, `insert`, `update` and `delete`.

Any changes that your commands make to Excel tables are immediately visible in Excel. 

The default connection is in-memory, so any objects you create in the session (tables, views) are temporary and will disappear as soon as you disconnect.

## The `__address` column
When using the SQLite engine, all workbook tables get an additional hidden column called `__address`. This column contains the original address of the row in Excel. Double-clicking the address in the results grid will select the row in Excel. The `__address` column is **not included** in the results if you only specify `*` in the select list; you must include it in the select list explicitly if you need it (e.g. `select *, __address...`).

## Working with cells
While the SQLite engine primarily works with Excel tables, you can also work with cells using the **`xlcells()`** table-valued function. 

The following query returns a list of cells in the current selection:
```sql
select * from xlcells()
``` 
We can also return a list of cells in a specified range, like so:
``` sql
select * from xlcells('Sheet1!B5:D9')
```
Here's what the returned data looks like:
![Cells query](../../Images/xlcells.png)

For each cell in the selection, one row is returned in the results. Each cell is described with the following attributes: `address`, `row`, `column`, `column letter`, `value`, `type`, `formula`.

You can use this information to search for values or format cells based on criteria that your specify with SQL.

Aside from reading values, the `xlcells` function can also be used for updating. In that case, `xlcells` is used as a table and its parameter is specified in the `where` clause. 

For example, the follwing query will add 100 to all cells that have a numeric value:

```sql
update
	XLCells -- referenced like a table
set 	
	Value = Value + 100
where 
	targetRangeAddress = 'H6:I8' -- the parameter is here (visible in autocomplete)
	and [Type] = 'Double'
```
> Under the hood, table-valued-functions in SQLite are implemented as virtual tables, and can be used both as tables and as tabular functions.


## Formatting rows and cells
Since the SQLite engine is running *in-process* with Excel, it can interact with Excel objects. A typical use case for this is modifying formatting. 

QueryStorm's SQLite engine offers the following functions for this purpose: `SetBackgroundColor` and `ClearBackgroundColor`. Both functions require the target address (`__address` in case of rows and `Address` in case of cells).

Here's an example:
``` SQL
--first clear any previous formatting from the 'movies' table (the table name is used as the address)
select ClearBackgroundColor('movies');

--apply new formatting to target rows
select
	*, SetBackgroundColor(__address, 'Orange')
from
	 movies
where
	gross > 400000000
order by
	__row asc
``` 
And the resulting formatting look like this:
![Formatting rows example](../../Images/setbackgroundcolor.png)



## Column indexes
All columns of Excel tables are automatically indexed by the SQLite engine. This makes joins and searches very fast, though all of the indexes are single-column indexes. 

For representing workbook tables QueryStorm uses SQLite's virtual table mechanism, meaning that the indexing logic is baked in and no additional indexes can be defined (for workbook tables). 

If different indexing is needed, you can create a copy of the table and add indexes to the copy, though this is very rarely required.