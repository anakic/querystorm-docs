# Formatting rows via C# #

> [Click here](https://www.querystorm.com/Downloads/Demos/salaries_sf.xlsx) to download the sample workbook.

## Formatting row-by-row

Each row object has a `GetRange()` method that we can use to interact with the range in Excel that contains the row data. As an example, let's change the background color of all rows in Excel where the `TotalPay` is higher than $500k:

``` C#
salaries
	.Where(s => s.TotalPay > 500000) //find rows
	.ForEach(s => s.GetRange().Interior.ColorIndex = 38) //modify formatting
```
Here's what the result looks like:

![Coloring rows via C#](https://i.imgur.com/31pFsPz.png)

The result is fine, but this way of formatting rows can be slow, especially if there is a large number of rows to format.

## Bulk row formatting
If we're applying the same formatting to many rows, we can get much better performance by using the `Format` method. 

```csharp
//bulk formatting (takes ~0.7s)
salaries
	.Where(s=>s.Id % 2 == 1)
	.Take(50000)
	.Format(rng => rng.Interior.ColorIndex = 32)
	
```
The `Format` method uses some optimization mechanisms to minimize the amount of calls to Excel, while still achieving the same result. 

This query is about 10x faster than formatting row-by-row. In this example, it colors 50k rows in **0.7s**. Row-by-row formatting would take **6.5s**.

In the previous example, we applied a color to one group of rows. We can also process multiple groups of rows in the same query. For example, let's apply a color based on *TotalPay* brackets: 

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
