# Working with cells
Aside from working with structured data, it's also possible to work with individual Excel cells. Two global methods are provided for this: `Cells()` and `Range()`.

The `Cells()` method returns a list of cells in a range. Here are some examples of how to use it:
- `Cells()` returns a list of cells in the current selection
- `Cells("A1:B2")` returns a list of four cells in the active sheet *(A1, A2, B1, B2)*
- `Cells("Sheet2!A1:B2")` same as above, but the cells belong to *Sheet2*
- `Cells("abc")` returns a list of cells in a named range or a table called *"abc"*

## Modifying cell values
In the following example, some cells do not have spaces between the first and last name of the person. I select the cells in Excel, and use the `Regex.Replace()` method to insert a space where a lowercase letter is immediately followed by an uppercase letter.

``` C#
Cells().Do(cell => cell.Value = Regex.Replace(cell.Value, "([a-z])(A-Z)", "$1 $2"))
```
> In the replacement string $1 and $2 are "substitutions" that refer to group captures (the two captured letters).

![Individual cells + Regex](https://i.imgur.com/rYm5BEZ.png)

## Modifying cell formatting
In the next example, I want to color all cells of the current selection that contain a value of type `string`:

``` C#
Cells()
	.Where(cell=>cell.Value?.GetType() == typeof(string))
	.Do(cell => cell.Interior.ColorIndex = 40)
```

![Cell coloring](https://i.imgur.com/nWJzEYZ.png)

## Cell values as parameters
Using cells as parameters can come in handy, especially when automating queries. Here's how to read the value from a single cell:

``` C#
var abc = (int)Range("Sheet1!A2").Value;
```
Or alternatively:
``` C#
var def = (string)Cells("Sheet1!B2").Single().Value;
```
> In a future version, named ranges will be available to the script as variables.