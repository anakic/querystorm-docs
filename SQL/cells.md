# Working with individual cells

Aside from working with tables, you can also work with cells using the **`xlcells()`** table-valued function.

The following query returns a list of cells in the current selection:
```sql
select * from xlcells()
``` 
We can also return a list of cells in a specified range, like so:
``` sql
select * from xlcells('Sheet1!B5:D9')
```
Here's what the returned data looks like:
![Cells query](https://i.imgur.com/iHcra7e.png)

For each cell in the selection, one row is returned in the results. Each cell is described with the following attributes: `address`, `row`, `column`, `column letter`, `value`, `type`, `formula`.

### Updating cell values

The primary reason to use the `xlcells()` function is to make changes to arbitrary selections of cells. In SQLite, table-valued functions are treated as tables, so we can use them in update statements. When updating `xlcells` only the `Value` column can be updated; updating other columns will have no effect. 

For example, let's reverse the text in the selected cells:
``` SQL
update
	XLCells
set
	Value = reverse(Value)
```
![Update selected cells](../images/cells_update1.png)

!!! Tip
	You can select multiple areas in Excel by holding down the `Ctrl` key while selecting them in Excel. To navigate between selected areas you can use the `Ctrl+Alt+Left` and `Ctrl+Alt+Right` shortcuts.

The `xlcells` function can also take a parameter: `targetRangeAddress`. Parameters are treated as (hidden) columns by SQLite, and can be supplied in the `where` clause when updating.

Here's how to reverse the text in a specified range:
```SQL
update
	XLCells
set
	Value = reverse(Value)
where
	targetRangeAddress = 'a1:c10'
``` 
![Update selected cells](../images/cells_update2.png)

