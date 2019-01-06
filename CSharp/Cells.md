# Working with cells
Aside from working with structured data, it's also possible to work with individual Excel cells. 

The `Cells()` method returns a **list of cells in a range**:

- `Cells()` returns a list of cells in the current selection
- `Cells("A1:B2")` returns a list of four cells in the active sheet *(A1, A2, B1, B2)*
- `Cells("Sheet2!A1:B2")` same as above, but the cells belong to *Sheet2*
- `Cells("abc")` returns a list of cells in a named range or a table called *"abc"* (including the header)

The returned objects are not Excel Range objects, but lightweight `CellAccessor` objects. These objects make reading, writing and formatting cells very fast and convenient. 

## Modifying cell values
In the following example, all spaces in the table have been removed. In order to return spaces between words I'm going to use regular expressions and introduce a space in places where a lowercase letter is immediately followed by an uppercase letter:

``` C#
Cells().UpdateValues(c => Regex.Replace(c.Value, "([a-z])([A-Z])", "$1 $2"))
```

![Update cells](../images/cells_update.png)

In the example, I'm using the `UpdateValues` extension method that is defined for `IEnumerable<CellAccessor>` objects. Since I did not supply an address to the `Cells` method, the update was carried out on the selected cells. 

## Modifying cell formatting
We can also format cells. Let's highlight all cells of the current selection that contain a value of type `string`:

``` C#
Cells()
    .Where(cell => cell.Type == typeof(string))
    .Format(range => range.Interior.ColorIndex = 40)
```

![Cell coloring](https://i.imgur.com/nWJzEYZ.png)

The `Format` method tries to group cells into rectangular ranges before calling the provided action on the ranges. This reduces the number of calls to Excel and improves performance since formatting is not done cell-by-cell.  