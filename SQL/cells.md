# Working with individual cells

Aside from working with tables, QueryStorm's SQLite engine can also work with cells, by using the **`cells()`** table-valued function.

The following query will return a list of cells in the current selection:
```sql
select * from cells()
``` 
We can also return a list of cells in a specified range, like so:
``` sql
select * from cells('Sheet1!B5:D9')
```
Here's what the returned data looks like:
![Cells query](https://i.imgur.com/iHcra7e.png)

For each cell in the selection, one row is returned in the results. Each cell is described with the following attributes: `address`, `row`, `column`, `column letter`, `value`, `type`, `formula`.

### Updating cell values

The primary use case for the `cells()` function is making changes to an arbitrary selection of cells (e.g. doing a regex replace operation, or conditionally formatting them). 

To achieve this, we need to be able to update the values in the cells. Currently, updating cells via a SQL `update` command is not supported (it might be in future versions), but we can update them using the `SetCellValue()` function.

For example, let's reverse the text in the selected cells:
``` SQL
SELECT
	*, SetCellValue(Address, reverse(Value))
FROM
	cells()
```
![Update cell values example](https://i.imgur.com/XbDgKfu.png)

Using a select statement to do an update might seem strange, but its fairly easy to get used to. In a future version, it's also likely that support for updating it via direct `UPDATE` statement will be added. 

!!! Tip
	The selection can consist of one or more areas. To select multiple areas, hold down the `Ctrl` key while selecting them in Excel. 
