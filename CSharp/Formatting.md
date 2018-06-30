# Formatting rows via C# #

> You can [download the sample workbook](https://www.querystorm.com/Downloads/Demos/salaries_sf.xlsx) if you'd like to follow along.

We can get the Excel range of any row in a table by using the `GetRange()` method. We can use this to modify row formatting.

As an example, let's change the background color of all rows in Excel where the `TotalPay` is higher than $500k:

``` C#
salaries
	.Where(s => s.TotalPay > 500000) //find rows
	.Do(s => s.GetRange().Interior.ColorIndex = 38) //modify formatting
```
The `Do` method is similar to `ForEach`, except that it executes on the main thread (making interactions with Excel faster). It is recommended to use `Do` instead of `ForEach` when interacting with a large number of Excel objects.

Here's what the result looks like:

![Coloring rows via C#](https://i.imgur.com/31pFsPz.png)

#### Bulk row formatting
With `Do()`, coloring rows is relatively fast, but its still coloring them row by row. If we're applying the same formatting to a huge number of rows, we can make things much faster by grouping the rows and applying formatting in bulk to each group. 

The following query applies a color to 50k rows that have an odd number as the Id:

```csharp
//row by row formatting (takes ~6.5s)
salaries
	.Where(s=>s.Id % 2 == 1)
	.Take(50000)
	.Do(s => s.GetRange().Interior.ColorIndex = 31)
```

This query colors 50k rows in a row-by-row fashion and takes **6.5s** to complete. 

We can apply formatting to a group of rows much faster using the `Format()` method which internally optimizes communication with Excel:

```csharp
//bulk formatting (takes ~0.7s)
salaries
	.Where(s=>s.Id % 2 == 1)
	.Take(50000)
	.Format(rng => rng.Interior.ColorIndex = 32)
	
```
This query achieves the same end result as the previous one, but only takes **0.7s** to complete! 

We can also process different groups with different colors. Let's color the rows based on a histogram of values in the `TotalPay` column, so that e.g. people who earn 0-20k get one color, 20k-40k another etc... Here's how to do it:

```csharp
int binSize = 20000;

Random rnd = new Random();

salaries
    .GroupBy(g => ((int)g.TotalPay / binSize) * binSize) //create bins
    .Select(g=> new { ColorIndex = rnd.Next(1,49), Group = g }) //define address groups and colors
    .Do(x=>x.Group.Format(r=>r.Interior.ColorIndex = x.ColorIndex)) //run coloring
``` 
And here's the result:

![Coloring based on histogram](https://i.imgur.com/slHylS5.png)

> The rows do not have to be adjacent in order to be in the same group, although it helps performance if they are.

#### Clear formatting
It can be useful to clear any left over formatting before applying new formatting. Here's how to clear coloring from all rows of a table:

``` C#
salaries.Range.Interior.ColorIndex = 0;
``` 
> Each table object has a Range property we can use to reference the range that contains the table. There is also a `ListObject` property we can use to directly reference the Excel ListObject.
