# Working with data (LINQ)
QueryStorm exposes tables as collections of strongly typed objects, allowing you to easily query and modify data inside them using C# and LINQ.

!!! Video
	In case you prefer video, check out this [introduction video](https://vimeo.com/242216594 "C# engine intro video on Vimeo").

!!! Data
	The examples in this document use a [public dataset about salaries](../demofiles/salaries_sf.xlsx "Salaries in San Francisco dataset") in the public sector in San Francisco.

## Let's get querying
To connect to the active workbook, click the **Connect (C#)** icon in the QueryStorm ribbon:

![Connect with C#](https://i.imgur.com/GYMfRJA.png)

Once connected, Excel tables show up in the object explorer and are available to your scripts as variables. Each table is a collection of strongly typed rows. The row types are generated dynamically. In this example, there's only one table: `salaries`.

![Connected with C#](https://i.imgur.com/FQoRvi1.png)

To demonstrate basic querying, let's find people who earn between 100k and 105k:

``` C#
salaries.Where(s=>s.TotalPay >= 100000 && s.TotalPay <= 105000)
``` 

Here are the results:

![Sample result](https://i.imgur.com/xAZdx0w.png)

For a slightly more complicated example, let's find the highest paid worker for each job title:

``` C#
salaries
	.GroupBy(s => s.JobTitle)
	.Select(g => g.OrderByDescending(s=>s.TotalPay).First())
```  

#### Funky column names
Table columns in Excel can have names that are not legal C# identifiers. Property names are based on the column headers but are normalized in order to be legal C# identifiers: non-alphanumeric characters are removed, names that start with a number are prefixed with an underscore.

For example, if the table has a column named `Job Title`, the corresponding property will be `JobTitle` (without space):

``` C#
salaries.GroupBy(s => s.JobTitle)
``` 

!!! Notes
	- In the results grid, column names will show up with their original names because the generated properties have a `[System.ComponentModel.DisplayName]` attribute applied.
	- You can also access properties via indexer, e.g. `row["Job Tilte"]`. The return type of the indexer is `dynamic`.  

#### Locating rows in Excel
In order to find a particular row in Excel, we can include the row address in the results. Each row has a `GetAddress()` method that can be used to find the row in Excel.

!!! Note
	`GetAddress()` is a method rather than a property solely to avoid a naming collision in cases where a table has a column named *Address*. 

Let's find people with the job title "Transit Operator" in Excel.

``` C#
salaries
	.Where(s => s.JobTitle == "Transit Operator")
	.Select(s => new { s.Id, s.EmployeeName, s.TotalPay, Address = s.GetAddress()})//include Address as a property
```

Once we have the address in the results, we can double-click the **address cell** OR the **row header** to navigate to the row in Excel.

![Select row in Excel](https://i.imgur.com/J9QxBIT.png)

!!! Note 
	If there is no address in the results, double-clicking the row header will have no effect.

### Updating rows
To modify data in an Excel table, we need to modify the rows and then commit the changes. Here's how to modify the `TotalPay` of all Transit Operators:

``` C#
//prepare changes
salaries
	.Where(s => s.JobTitle == "Transit Operator")
	.ForEach(s=> s.TotalPay += 10000);
	
//commit changes into Excel table
salaries.SaveChanges();
```

Note that changes must be explicitly saved in order to be committed into the Excel table. We can do this in a single command, though, by using the `Update()` method:

``` C#
salaries
	.Where(s => s.JobTitle == "Transit Operator")
	.Update(s=> s.TotalPay += 10000)
```

### Deleting rows
We can also delete rows. Let's delete rows where the TotalPay is zero:

``` C#
salaries.Where(s => s.TotalPay == 0).Delete()
```
 
### Inserting rows
Finally, we can also insert rows:

``` C#
salaries.Insert(99999, "John Smith", "Journalist", 0, 0, 0, 0, 99999, 0, 2017, null, "San Francisco", "PT");
```

The `Insert` method is strongly typed, but all of its arguments are optional, allowing you to skip the ones where `null` should be inserted (as well as calculated columns).

Saving changes to Excel can be slow, which can be an issue when inserting thousands of rows. In such cases, inserts should be staged via the `InsertIntoCache()` method and finally committed to Excel by a single call to `SaveChanges()`:

```csharp
people.InsertIntoCache(10, "abc");
people.InsertIntoCache(11, "def");
people.InsertIntoCache(12, "ghi");
people.SaveChanges(); //saves the new rows into Excel
```

## C# scripting syntax

The scripting flavor of C# is supported by Roslyn (the C# compiler) and is slightly different from regular C#. Most *normal* C# syntax is also valid in scripts, but scripts also allow a more relaxed syntax where you can evaluate expressions, without all the ceremony of defining types and methods. 

!!! Note
	Under the hood, Roslyn preprocesses scripts and turns them into regular C# before compiling to IL.

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

Another thing scripts can do is reference other scripts and libraries. This is supported via `#r` and `#load` directives that are placed at the beginning of scripts:

```csharp
#r "System.IO.Compression"

//...code that uses the System.IO.Compression assembly...
```
```csharp
#load "anotherScript"

//...code that uses members from "anotherScript"...
```
These expansions to the C# syntax are provided by Roslyn (not QueryStorm). For more on referencing scripts and libraries, click [here](./references).