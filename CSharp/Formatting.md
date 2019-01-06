# Formatting rows via C# #

!!! Data
	Download the sample workbook from [here](https://www.querystorm.com/Downloads/Demos/salaries_sf.xlsx).

## Formatting row-by-row

Each row object has a `GetRange()` method that we can use to interact with it in Excel. Let's use it to change the background color of all rows in Excel where the `TotalPay` is higher than $500k:

``` C#
salaries
	.Where(s => s.TotalPay > 500000) //find rows
	.ForEach(s => s.GetRange().Interior.ColorIndex = 38) //modify formatting
```
Here's what the result looks like:

![Coloring rows via C#](https://i.imgur.com/31pFsPz.png)

The result is fine, but this way of formatting rows (one-by-one) can be slow, especially if there are thousands of rows to format.

## Bulk row formatting
If we're applying the same formatting to many rows, we can get much better performance by using the `Format` extension method that is defined for `IEnumerable<TRow>` objects. 

```csharp
//bulk formatting (takes ~0.7s for 50k rows)
salaries
	.Where(s => s.Id % 2 == 1)
	.Take(50000)
	.Format(rng => rng.Interior.ColorIndex = 32)
	
```
The `Format` method uses several optimizations to minimize the number of calls to Excel, while still achieving the same result. This makes it about 10x faster compared to row-by-row formatting (in this example ~0.7s instead of ~6.5s for 50k rows).

In the previous example, we applied a color to a single group of rows. We can also process multiple groups of rows in the same query. For example, let's apply a color based on *TotalPay* brackets: 

```csharp
int binSize = 20000;

Random rnd = new Random();

salaries
    .GroupBy(g => ((int)g.TotalPay / binSize) * binSize) //create bins
    .ForEach(g=> g.Format(r=>r.Interior.ColorIndex = rnd.Next(1,49))) //apply a color to each group
``` 
And here's the result:

![Bulk formatting histogram](../images/bulk_formatting_histogram.png)

#### Clear formatting

Each table object has a `Range` property that we can use to access the entire range of the table in Excel and clear any left over formatting before applying new formatting. Here's how to clear coloring from all rows of a table:

``` C#
salaries.Range.Interior.ColorIndex = 0;
``` 
