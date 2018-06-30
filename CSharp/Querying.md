# Working with data (LINQ)
Excel has a concept of tables ([not to be confused with sheets](http://www.contextures.com/xlExcelTable01.html "Excel tables overview")) that represent structured data, similar to database tables or collections in .NET. QueryStorm exposes tables as collections of strongly typed objects, allowing you to easily query and modify data inside them using C# and LINQ.

> **Demo workbook**
> The examples in this document use a dataset about salaries in the public sector in San Francisco. Please **[download it](https://www.querystorm.com/downloads/salaries_sf.xlsx "Salaries in San Francisco dataset")** to follow along.

### Let's get querying
To connect to the active workbook, click the **Connect (C#)** icon in the QueryStorm ribbon.

![Connect with C#](https://i.imgur.com/GYMfRJA.png)

Once connected, Excel tables show up in the object explorer and are available to your scripts as variables. Each table is a collection of strongly typed rows. The row types are generated dynamically. In this example, there's only one table, and it's called `salaries`.

![Connected with C#](https://i.imgur.com/FQoRvi1.png)

To demonstrate basic querying, let's find people who earn between 100k and 105k:

``` C#
salaries.Where(s=>s.TotalPay >= 100000 && s.TotalPay <= 105000)
``` 

Here's the result:

![Sample result](https://i.imgur.com/xAZdx0w.png)

Now let's find the highest paid worker for each JobTitle:

``` C#
salaries
	.GroupBy(s => s.JobTitle)
	.Select(g => g.OrderByDescending(s=>s.TotalPay).First())
```  
Piece of cake.

#### Funky column names
Table columns in Excel can have names that are not legal C# identifiers. It's (currently) not possible to reference them as properties, but we can access them via indexer. 

For example, if the table had a column named `Job title` (with a space) we could reference it like so:

``` C#
salaries.GroupBy(s => s["Job title"])
``` 

> The return type of the indexer is `dynamic` so you might want to cast it to the expected type. 

#### Locating rows in Excel
In order to find a particular row in Excel, we can include the row address in the results. Each row has a `GetAddress()` method that can be used to find the row in Excel.

Let's find people with the job title of "Transit Operators" in Excel.

``` C#
salaries
	.Where(s => s.JobTitle == "Transit Operator")
	.Select(s => new { s.Id, s.EmployeeName, s.TotalPay, Address = s.GetAddress()})//include Address as a property
```

Once we have the address in the results, we can double-click the **address cell** OR the **row header** to navigate to the row in Excel.

![Select row in Excel](https://i.imgur.com/J9QxBIT.png)

> Note #1: If there is no address in the results, double-clicking the row header will have no effect.
> Note #2: Creating a projection (`new {...}`) with many columns just to include the address can be tedious. A workaround is in the works.

### Updating rows
To modify data in an Excel table, we need to modify the row objects and then commit the changes.

Here's how to raise the `TotalPay` of all Transit Operators by 10k:

``` C#
//prepare changes
salaries
	.Where(s => s.JobTitle == "Transit Operator")
	.ForEach(s=> s.TotalPay += 10000);
	
//commit changes into Excel table
salaries.SaveChanges();
```

Note that **changes must be explicitly committed** in order to be propagated into the Excel table. To do this as a single expression, we can use the `Update()` method:

``` C#
salaries
	.Where(s => s.JobTitle == "Transit Operator")
	.Update(s=> s.TotalPay += 10000)
```

### Deleting rows
We can also delete rows. Let's delete rows where the TotalPay is zero:

``` C#
salaries
	.Where(s => s.TotalPay == 0)
	.Delete()
```
 
### Inserting rows
Finally, we can also insert rows. Let's insert a row into the `salaries` table:

``` C#
salaries.Insert(99999, "John Smith", "Journalist", 0, 0, 0, 0, 99999, 0, 2017, null, "San Francisco", "PT");
```

The `Insert` method is strongly typed, but all its arguments are optional, allowing you to skip the ones where null should be inserted (as well as calculated columns).

> There is also an `InsertIntoCache` method which will insert data into the cache without committing the changes to Excel. Saving changes to Excel can be slow, especially when inserting thousands of rows. In such cases, all inserts should be staged using multiple calls to `InsertIntoCache` followed by a single call to `SaveChanges()`.  

