# Working with data (LINQ)
Excel has a concept of tables ([not to be confused with sheets](http://www.contextures.com/xlExcelTable01.html "Excel tables overview")) that represent structured data, similar to database tables or .NET collections. QueryStorm exposes tables as collections of strongly typed objects, allowing you to easily query and modify data inside them using C# and LINQ.

In case you prefer video to text, here's a video introduction to using C# in Excel:
[![C# engine - intro video](https://i.imgur.com/FvWhnu3.png)](https://vimeo.com/242216594 "C# engine intro video on Vimeo")

> **Demo workbook**
> The examples in this document use a dataset about salaries in the public sector in San Francisco. Please **[download it](../demofiles/salaries_sf.xlsx "Salaries in San Francisco dataset")** to follow along.

## Let's get querying
To connect to the active workbook, click the **Connect (C#)** icon in the QueryStorm ribbon:

![Connect with C#](https://i.imgur.com/GYMfRJA.png)

Once connected, Excel tables show up in the object explorer and are available to your scripts as variables. Each table is a collection of strongly typed rows. The row types are generated dynamically. In this example, there's only one table, and it's called `salaries`.

![Connected with C#](https://i.imgur.com/FQoRvi1.png)

To demonstrate basic querying, let's find people who earn between 100k and 105k:

``` C#
salaries.Where(s=>s.TotalPay >= 100000 && s.TotalPay <= 105000)
``` 

Here are the results:

![Sample result](https://i.imgur.com/xAZdx0w.png)

Now let's find the highest paid worker for each JobTitle:

``` C#
salaries
	.GroupBy(s => s.JobTitle)
	.Select(g => g.OrderByDescending(s=>s.TotalPay).First())
```  

#### Funky column names
Table columns in Excel can have names that are not legal C# identifiers. Such columns will be represented by normalized property names where non-alphanumeric characters are removed.

For example, if the table has a column named `Job title` (with a space) we could reference it like so:

``` C#
salaries.GroupBy(s => s.JobTitle)//space was removed
``` 

In the results grid, however, column names will show up with their original names. How can this be? The properties on the dynamically generated row type have the `[System.ComponentModel.DisplayName]` attribute applied. The results grid and the table writer check for this attribute so you can use it to control how headers for your own types look in the results grid, and when written into the workbook.  

!!! Note
	You can also access the value of any property in a row via indexer, e.g. `row["Job Tilte"]`. The indexer returns an object of type `dynamic`.  

#### Locating rows in Excel
In order to find a particular row in Excel, we can include the row address in the results. Each row has a `GetAddress()` method that can be used to find the row in Excel.

!!! Note
	The reason `GetAddress()` is a method rather than a property is simply to avoid a naming collision if the table has a column named *Address* (which is not unlikely). 

Let's find people with the job title "Transit Operator" in Excel.

``` C#
salaries
	.Where(s => s.JobTitle == "Transit Operator")
	.Select(s => new { s.Id, s.EmployeeName, s.TotalPay, Address = s.GetAddress()})//include Address as a property
```

Once we have the address in the results, we can double-click the **address cell** OR the **row header** to navigate to the row in Excel.

![Select row in Excel](https://i.imgur.com/J9QxBIT.png)

!!! Note 
	- If there is no address in the results, double-clicking the row header will have no effect.
	- Creating a projection (`new {...}`) with many columns just to include the address can be tedious. A workaround is in the works.

### Updating rows
To modify data in an Excel table, we need to modify the row objects and then commit the changes. Here's how to modify the `TotalPay` of all Transit Operators:

``` C#
//prepare changes
salaries
	.Where(s => s.JobTitle == "Transit Operator")
	.ForEach(s=> s.TotalPay += 10000);
	
//commit changes into Excel table
salaries.SaveChanges();
```

Note that **changes must be explicitly committed** in order to be propagated into the Excel table. To do this in a single command, we can use the `Update()` method:

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

The `Insert` method is strongly typed, but all its arguments are optional, allowing you to skip the ones where `null` should be inserted (as well as calculated columns).

There is also an `InsertIntoCache` method which will insert data into the cache without committing the changes to Excel. Saving changes to Excel can be slow, especially when inserting thousands of rows. In such cases, all inserts should be staged via calls to `InsertIntoCache` and finally committed by a single call to `SaveChanges()`:

```csharp
people.InsertIntoCache(10, "abc");
people.InsertIntoCache(11, "def");
people.InsertIntoCache(12, "ghi");
people.SaveChanges(); //saves the new rows into Excel
```

## C# scripting syntax

The scripting flavor of C# is supported by Roslyn (the C# compiler) and is slightly different from regular C#. Regular C# code is mostly valid in scripts, but scripts also allow a more relaxed syntax where you can evaluate expressions without all the ceremony of defining types and methods. 

!!! Note
	Under the hood, Roslyn preprocesses scripts and turns them into regular C#.

For example, we can return the current date like so:
```csharp
DateTime.Now
``` 
We didn't need to define a class and `Main` method. We didn't need an explicit return statement, either. The last statement in the script is assumed to be an expression that needs to be evaluated and returned, unless it is terminated with a semicolon. **Closing an expression with a semicolon, however, hides the output**.

If you want to return a value from a statement that is not the last statement in the script, you need to explicitly return it with a `return` statement, e.g.:

```csharp
return Add(1,2);

public int Add(int a, int b)
{
	return a + b;
}

```

Another things scripts can do is reference other scripts and libraries. This is supported via `#r` and `#load` directives that are placed at the beginning of scripts:

```csharp
#r "System.IO.Compression"

//...code that uses the System.IO.Compression assembly...
```
```csharp
#load "anotherScript"

//...code that uses members from "anotherScript"...
```
For more on referencing scripts and libraries, click [here](./references).