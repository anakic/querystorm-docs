# C# scripting in Excel

Being a modern language with access to a rich ecosystem of libraries, C# can be immensely helpful both for automation and data processing in Excel. The goal of QueryStorm in this regard is to make C# easily available in Excel in order to help developers, db professionals and data scientists make better use of Excel. The C# scripting engine in QueryStorm is powered by the Roslyn compiler.

In this article, I'm going to introduce QueryStorm's C# scripting capabilities and give you some examples of how to use it. But first - an intro video:

[![C# engine - intro video](https://i.imgur.com/FvWhnu3.png)](https://vimeo.com/242216594 "C# engine intro video on Vimeo")


> **Demo workbook**
> The examples in this document use a dataset about worker salaries in San Francisco. Please [download it](https://www.querystorm.com/downloads/salaries_sf.xlsx "Salaries in San Francisco dataset") if you want to follow along.

> **New to QueryStorm?**
> You can download QueryStorm and get a trial key by clicking the *"Try QueryStorm"* button on the [QueryStorm homepage](https://www.querystorm.com "QueryStorm homepage"). If you have any issues installing or running QueryStorm, please [get in touch](mailto:antonio@querystorm.com?subject=Troubleshooting "get in touch").

> **Do you need a QueryStorm license?** 
> You can use QueryStorm's C# engine without a license, but in order to get intellisense support, error highlighting, auto-formatting and other premium features, you'll need to unlock QueryStorm. You can do this with a free trial or by [purchasing a full QueryStorm license](https://www.querystorm.com/#pricing "Buy QueryStorm license").
>
 
## Querying and modifying data
Excel doesn't impose strict limitations on data entry, but it does have a concept of tables ([not to be confused with sheets](http://www.excel-easy.com/data-analysis/tables.html "Excel tables intro")). **Excel tables represent structured data**, similar to database tables or collections in .NET. With its LINQ capabilities, C# is well suited for dealing with structured and relational data, and QueryStorm makes using C# in Excel very convenient. 

### Let's get querying
To connect to the active workbook and open the IDE, click the **Connect (C#)** icon in the QueryStorm ribbon. 

![Connect with C#](https://i.imgur.com/GYMfRJA.png)

Once connected, all Excel tables show up in the object explorer and are made available to user scripts as variables. Each table is a collection of strongly typed rows. The row types are generated dynamically by QueryStorm. In this example, there's only one table, and it's called `salaries`.

![Connected with C#](https://i.imgur.com/FQoRvi1.png)

To demonstrate basic querying, let's find people who earn between 100k and 105k:

``` C#
salaries.Where(s=>s.TotalPay >= 100000 && s.TotalPay <= 105000)
``` 

This query produces a result that looks like this:

![Sample result](https://i.imgur.com/xAZdx0w.png)

Now let's find the highest paid worker for each JobTitle:

``` C#
salaries
	.GroupBy(s => s.JobTitle)
	.Select(g => g.OrderByDescending(s=>s.TotalPay).First())
```  
With plain Excel this would be tricky, but with LINQ it's a piece of cake!

#### Funky column names
Table columns in Excel can have names that are not legal C# identifiers. It's not possible to reference these properties directly in C#, but we can access them via indexer. 

For example, if the table had a column named `Job title` (with a space) we could reference it like so:

``` C#
salaries.GroupBy(s => s["Job title"])
``` 
> The return type of the indexer is `dynamic` so you might want to cast it to the expected type. 

#### Locating rows in Excel
In order to find a particular row in Excel, we can include the row address in the results. Each row has a `GetAddress()` method that can be used to find the row in Excel.

Let's look for all transit operators and make sure we include the address so we can find them in Excel.

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

Here's how to raise the TotalPay of all Transit Operators by 10k:

``` C#
//prepare changes
salaries
	.Where(s => s.JobTitle == "Transit Operator")
	.ForEach(s=> s.TotalPay += 10000);
	
//commit changes into Excel table
salaries.SaveChanges();
```

Note that **changes must be explicitly committed** in order to be propagated into the Excel table. To do this as a single expression, we can use the `Do()` method which is similar to `ForEach()`, but automatically commits changes before returning:

``` C#
salaries
	.Where(s => s.JobTitle == "Transit Operator")
	.Do(s=> s.TotalPay += 10000)
```

> Use the `Do()` method when ever you are making changes to a collection of rows and you want to immediately commit them. If you intend to make additional changes before committing, use the `ForEach()` extension method or a loop instead.

### Deleting rows
We can also delete rows. Let's delete rows where the TotalPay is zero:

``` C#
salaries
	.Where(s => s.TotalPay == 0)
	.Do(s=> s.Delete())
```

> The `Delete()` method marks the row for deletion. As with updates, the changes need to be committed in order to be propagated to Excel. The `Do()` method does that internally for us.
 
### Inserting rows
Finally, we can also insert rows. Let's insert a row into the `salaries` table:

``` C#
salaries.Insert(99999, "John Smith", "Journalist", 0, 0, 0, 0, 99999, 0, 2017, null, "San Francisco", "PT");
salaries.SaveChanges();
```
> As before, changes need to be explicitly committed in order to be propagated to the Excel table.

## Formatting rows
We can get the Excel range of any row by using the `GetRange()` method. We can use this to modify row formatting.

As an example, let's change the background color of all rows in Excel where the `TotalPay` is higher than $500k:

``` C#
salaries
	.Where(s => s.TotalPay > 500000) //find rows
	.Do(s => s.GetRange().Interior.ColorIndex = 38) //modify formatting
```

Here's what the result looks like:

![Coloring rows via C#](https://i.imgur.com/31pFsPz.png)

Notice that I'm using the `Do()` method again, even though I'm not committing any data changes to an Excel table. Aside from auto-committing changes to tables, the `Do()` method has one more special feature: **it pushes actions to the main thread**. This is important because the C# engine executes scripts on a background thread, but **working with ranges on the main thread is much faster** (due to Excel's internal reasons).

On my machine, coloring 1k rows takes:
- **0.3s** with `Do()` 
- **4s** with `ForEach()`

> Note #1: **Use `Do()` instead of `ForEach()` when ever manipulating Excel objects!**
> Note #2: There is currently no intellisense support for `Range` objects. This is a technical issue and, hopefully, temporary. 

### Bulk row formatting
With `Do()`, coloring rows is relatively fast, but its still coloring them row by row. If we're applying the same formatting to a huge number of rows, we can make things much faster by grouping the rows and applying formatting in bulk. 

Consider the following query that applies a color to 50k rows that have an odd number as the Id:

``` C#
//row by row formatting (takes ~6.5s)
salaries
	.Where(s=>s.Id % 2 == 1)
	.Take(50000)
	.Do(s => s.GetRange().Interior.ColorIndex = 31)
```

This query colors 50k rows in a row-by-row fashion and takes **6.5s** to complete. 

Since we're applying the same formatting to all target rows, we can get much better performance if we apply formatting in bulk:

``` C#
//bulk formatting (takes ~0.7s)
salaries
	.Where(s=>s.Id % 2 == 1)
	.Take(50000)
	.ToAddressGroup()
	.Process(rng => rng.Interior.ColorIndex = 32)
	
```

This query achieves the same end result as the previous one, but only takes **0.7s** to complete!  

> `ToAddressGroup()` is an extension method on `IEnumerable<row>`. It returns an object for bulk formatting an entire  group of rows.

We can also process different groups with different colors. Let's color the rows based on a histogram of values in the `TotalPay` column, so that e.g. people who earn 0-20k get one color, 20k-40k another etc... Here's how to do it:

``` C#
int binSize = 20000;

Random rnd = new Random();

salaries
    .GroupBy(g => ((int)g.TotalPay / binSize) * binSize) //create bins
    .Select(g=> new { ColorIndex = rnd.Next(1,49), AddressGroup = g.ToAddressGroup() }) //define address groups and colors
    .Do(x=>x.AddressGroup.Process(r=>r.Interior.ColorIndex = x.ColorIndex)) //run coloring
``` 
And here's the result:

![Coloring based on histogram](https://i.imgur.com/slHylS5.png)

> The rows do not have to be adjacent in order to be in the same group, although it helps performance if they are.

### Clear formatting
It can be useful to clear any left over formatting before applying new formatting. Here's how to clear coloring from all rows of a table:

``` C#
salaries.Range.Interior.ColorIndex = 0;
``` 
> Each table object has a Range property we can use for this purpose.

## Working with individual cells
Aside from working with structured data, it's also possible to work with individual Excel cells. Two global methods are provided for this: `Cells()` and `Range()`.

The `Cells()` method returns a list of cells in a range. Here are some examples of how to use it:
- `Cells()` returns a list of cells in the current selection
- `Cells("A1:B2")` returns a list of four cells in the active sheet *(A1, A2, B1, B2)*
- `Cells("Sheet2!A1:B2")` same as above, but the cells belong to *Sheet2*
- `Cells("abc")` returns a list of cells in a named range or a table called *"abc"*

### Modifying cell values
In the following example, some cells do not have spaces between the first and last name of the person. I select the cells in Excel, and use the `Regex.Replace()` method to insert a space where a lowercase letter is immediately followed by an uppercase letter.

``` C#
Cells().Do(cell => cell.Value = Regex.Replace(cell.Value, "([a-z])(A-Z)", "$1 $2"))
```
> In the replacement string $1 and $2 are "substitutions" that refer to group captures (the two captured letters).

![Individual cells + Regex](https://i.imgur.com/rYm5BEZ.png)

### Modifying cell formatting
In the next example, I want to color all cells of the current selection that contain a value of type `string`:

``` C#
Cells()
	.Where(cell=>cell.Value?.GetType() == typeof(string))
	.Do(cell => cell.Interior.ColorIndex = 40)
```

![Cell coloring](https://i.imgur.com/nWJzEYZ.png)

### Cell values as parameters
Using cells as parameters can come in handy, especially when automating queries. Here's how to read the value from a single cell:

``` C#
var abc = (int)Range("Sheet1!A2").Value;
```
Or alternatively:
``` C#
var def = (string)Cells("Sheet1!B2").Single().Value;
```

## Managing references 
### Referencing dll's
QueryStorm's C# engine uses a standard set of *references* and *usings*. You can, however, add your own usings and references. 

For example, the `System.Windows.Forms` assembly isn't referenced by default. Here's how you can include it and show a message box to the user:

``` C#
#r "System.Windows.Forms" //add reference
using System.Windows.Forms; //include namespace 

MessageBox.Show("Hello world")

```  
The `#r` directive adds a reference to the specified assembly. 

Here are some examples of how to use it:
- `#r "System.Windows.Forms"` reference a named assembly from the GAC
- `#r "C:\packages\lib.dll"` reference a dll by absolute path
- `#r "\\serverA\packages\lib.dll"` reference a dll on a network share

> Note #1: Support for NuGet is planned, so something like `#r "nuget!some.package"` would add a reference to a NuGet package.
> Note #2: You can also customize the included references and usings by creating a custom connection via the *"Connect with"* button. 

### Referencing scripts
We can also load C# scripts in a similar way, using the `#load` directive. Here are some examples: 
- `#load "c:\mycode\class123.csx"` load a script form a local file
- `#load "\\server\scripts\demo.csx"` load a script form a network share
- `#load "my script"` load an embedded script from the current workbook
 
The `#load` directive accepts absolute paths and network shares, but it also accepts embedded scripts, as shown below:

![Referencing embedded script](https://i.imgur.com/RHSRZpg.png)

In the example above, I have a class called `MyCustomClass` defined in an embedded script called *my script*. Once I load the script, I can access the `MyCustomClass` type contained inside it, and I also get intellisense support for it.


## Multiple workbooks
As with all QueryStorm engines, the C# engine also supports working with tables from multiple workbooks. This can be achieved in one of two ways:

**A)** Via the dropdown in the ribbon

![Connect to all workbooks](https://i.imgur.com/5KruODX.png)


**B)** Via custom connection

![Select tables](https://i.imgur.com/o9lorUF.png)

## Loading external data
Being part of the .NET ecosystem, the C# engine has access to a large set of libraries, making it much easier to process data in various ways and retrieve data from various sources.

### Example #1: Getting data from a REST API
The QueryStorm website has a [REST API](https://www.querystorm.com/api/versions "versions API") for information on [versions](https://www.querystorm.com/versions.html "versions page"). In this exmple, my script calls this API, and writes the retrieved data into an Excel table:

``` C#
using Newtonsoft.Json;
using System.Net;

var response = await WebRequest.Create("https://www.querystorm.com/api/versions").GetResponseAsync();
string jsonStr = new StreamReader(response.GetResponseStream()).ReadToEnd();

//deserialize into objects
var items = JsonConvert.DeserializeObject<Item[]>(jsonStr);

//write results into Excel table
items.IntoTable("QueryStormVersions", true, "B2");

class Item
{
    public DateTime CreatedOn { get; set; }
    public Version Version { get; set; }
    public long Size { get; set; }
    //public string Hash { get; set; }
    //public string WhatsNew { get; set; }
    public string Url { get; set; }
}

```
> As can be seen in the example, scripts can use `async` features.

And here is the result:

![REST Example result](https://i.imgur.com/dlSYbrB.png)

Data can be written into ranges, tables and the clipboard. This is accomplished using the `IntoTable`, `IntoRange` and `IntoClipboard` extension methods (available on type `object`). The arguments to these methods correspond to the settings in the *Write results* dialog. The above example uses the `IntoTable(tableName, overwrite, defaultAddress)` method to write the data as an Excel table.

### Example #2: Simple RSS reader application
In the following video, I demonstrate building an RSS reader application in Excel. I think a video can do a better job than text of demonstrating what the procedure looks like, so here goes:

[![Screencast](https://i.imgur.com/6FtIOqp.png)](https://vimeo.com/241888977 "RSS Reader demo on Vimeo")


### Example #3: Getting data from a database
QueryStorm comes with its own database engine (SQLite) and support for connecting to external databases (SQL Server, MySql, Postgres, Access, ODBC). This makes it quite easy to fetch data from databases.

Once prepared, queries can be called just like functions, and they return data as `IEnumerable<dynamic>` objects.  

Here's an example...

**Step 1:** Connect to SQL Server: 

```
Server=mssql6.mojsite.com,1555; Database=thingieq_AdventureWorks2014; User Id=thingieq_abc; Password=D8R7VXLcSzss
```
**Step 2:** Prepare the SQL query, and embed it into the workbook e.g. as *"Products query"*:

``` SQL
SELECT * FROM Production.Product
```
**Step 3:** Call the query from C#

```C#
//execute query to fetch data from the database
var result = await Query("Products query").RunAsync();

//items collection is of type IEnumerable<dynamic>
var data = result.Items;

//Output data into Excel table
data.IntoTable("Parts", "B2")
```

Here's the result:
![running db query](https://i.imgur.com/mG7Hrrq.png)


## Under the hood
### Dynamically generated types
For a given Excel table, QueryStorm will dynamically generate a type (using `Reflection.Emit`) that will represent a table's row, in order to expose columns as strongly typed properties. This is done for each table in scope. Row objects don't have a copy of the row data, they only contain the index of the row in the table. Each property *knows* which column index it refers to, so it uses the row index and column index to access the correct value from the table's cache. This makes row objects extremely lightweight (they only contain a single integer).

### The cache
So what's the table cache? QueryStorm internally maintains an 2d-array (`object[,]`) for each table. This array is the table cache. The cache is selectively updated as changes are made to the Excel table (e.g. user changes a value). If the cache becomes corrupt for any reason, the user can explicitly reload it from the context menu of the table in the object explorer (this shouldn't really happen but...). Anyway, since the cache is entirely in memory, queries are very fast; e.g. ~30ms for reading 100k rows with 10 columns. Updating a row property writes the new value to the cache, and the cache keeps track of which regions have been changed. When the user calls `SaveChanges()`, only the changed parts of the cache are written into the Excel table, making updates very fast as well.   

### Roslyn
The execution of the scripts as well as the IDE features (intellisense, error squigglies, code auto-formatting) are powered by the Roslyn compiler. Each time the script text changes (user types a character), the code is analyzed and the editor updated to display any errors, intellisense suggestions, function documentation etc... Upon execution, a new thread is started in order to not block the UI and to enable the user to cancel execution. The version of C# that's made available to the user is C# 7, i.e. the latest version.  


### Licensing infrastructure
QueryStorm uses a licensing infrastructure called Keystodian. Keystodian is a separate project built in-house. It's not part of QueryStorm, but has not yet been offered commercially so at the moment only QueryStorm uses it. It revolves around the concept of digitally signed licenses and keys. Each key allows the owner to create one or more licenses. From the user's perspective a key is a GUID, and a license is a digitally signed XML document. 

Keystodian offers the following features:
- issuing and checking licenses 
- heartbeat with grace period (enables black-listing and license transfers)
- ad-hoc license transfers (users can easily transfer a license from machine to machine)
- trial licenses
- feature toggles
- payment integration (only [FastSpring](https://fastspring.com/ "FastSpring") so far)
- activating multiple licenses with a single key
- license and key expiration
- version capping
- admin and user dashboard

If you're interested in using **Keystodian** in your own products, please [get in touch](mailto:support@querystorm.com?subject=Keystodian "get in touch")!

## Planned improvements
- IDE features: debugger, refactorings, symbol look-up...
- Intellisense for Excel's objects
- An `IQueryable` provider for indexing. The indexing mechanism already exists (built for QueryStorm's SQLite engine), but it isn't yet being used by the C# engine.
- Ability to reference NuGet packages via the `#r` directive

## Acknowledgments
The C# scripting engine in QueryStorm uses the Roslyn compiler for execution and IDE support. Roslyn has a bit of a learning curve, but these excellent resources made things much easier:
- [Roslynpad](https://github.com/aelij/RoslynPad "RoslynPad") by Aelij
- Josh Varty's [Roslyn series](https://joshvarty.wordpress.com/learn-roslyn-now/ "Josh Varty's Roslyn blog")
- [Roslyn docs](https://github.com/dotnet/roslyn "Roslyn docs") on Github

QueryStorm also uses numerous other open source libraries, most notably [AvalonEdit](http://avalonedit.net/ "AvalonEdit") and [SQLite](http://sqlite.org/ "SQLite").
