# APIs available to scripts

In addition to the full Excel API that's available to user scripts, QueryStorm objects also expose a set of APIs that your scripts can interact with. The members and APIs that are available to user scripts are described below.

## Excel table APIs
Each Excel table is represented by a strongly typed object specific to that table. The following is a list of the relevant members that each table object exposes:

### Range
Gets the Range that contains the table.
```csharp
//"people" is an Excel table in this example
people.Range.Interior.ColorIndex = 0;
```

### ListObject
Gets the ListObject that the table object refers to.
```csharp
people.ListObject.ShowTotals = true;
```

### SelectedRows
Returns a list of rows that are part of the current selection in Excel.
```csharp
people.SelectedRows.FirstOrDefault()?.FirstName
```

### SaveChanges()
Commits any changes (updates, inserts and deletes) from the cache into the Excel table.
```csharp
people.First().Name ="abc"; 
people.SaveChanges();
```

### Insert(...)
Inserts the specified values into the table as a new row. Immediately commits the change into Excel.

```csharp
people.Insert(10, "abc");
```

!!! Note
	Since the `Insert` method is defined on the table, you might expect that `Update` and `Delete` methods would be as well. However, since they can work on arbitrary collections of rows, they are defined as extension methods on `IEnumerable<TRow>` instead. 

### InsertIntoCache(...)
Inserts the specified values into the table cache as a new row but does not commit to Excel. Useful for inserting a large number of rows (saving changes to Excel is slow).

```csharp 
people.InsertIntoCache(10, "abc");
people.InsertIntoCache(11, "def");
people.InsertIntoCache(12, "ghi");
people.SaveChanges(); //saves the new rows into Excel
```

## Named ranges
Each named range is represented by a global property that provides strongly typed access to the value contained in the named range. 

Suppose we have a named range *abc* that contains a numeric value. Here's how to access it:

```csharp
//reading the value in the named range "abc"
return abc + 5;
```

```csharp
//setting the value in the named range "abc"
abc = 123;
``` 

## Global methods
### Application
Returns the current Excel application.

```csharp
//example
Application().Workbooks.Open(@"c:\workbook.xls");
```

!!! Note
	`Application()` is a method rather than a property simply to avoid potential naming collisions with any table/range names present in the workbook.

### Cells(...)
Gets a collection of cells in the current selection or a specified range.
```csharp
//format the currently selected ranges
Cells().Format(r => r.Interior.ColorIndex = 34);

//format cells in range "a1:b2"
Cells("a1:b2").Format(r=>r.Interior.ColorIndex = 27);
```

### Csv(...)
Opens the specified CSV files and returns a list of rows, mapping the csv headers to the properties of the specified class.

```csharp

class Person
{
	public string Name { get; set; }

	//mapping the BirthDate property to header "Date of birth"
	[CsvHelper.Configuration.Attributes.Name("Date of birth")]
	public DateTime BirthDate { get; set; }

	public double Height { get; set; }
}

//Get rows in the people.csv file, mapped to the Person class.
IEnumerable<Person> people1 = Csv<Person>(@"c:\people.csv");

//...also specify delimiter
IEnumerable<Person> people2 = Csv<Person>(@"c:\people2.csv", "\t");

//...also specify datetime/number formats
IEnumerable<Person> people3 = Csv<Person>(@"c:\people2.csv", "\t", System.Globalization.CultureInfo.GetCultureInfo("pl"));
```

### Log(...)
Writes the specified value into the QueryStorm messages log. Drills down through properties to the specified depth. For debugging purposes. 

```csharp
//write the value of the person object into the messages pane, drill down into properties up to depth 2 
Log(person, 2);
```

### Query(...)
Gets an embedded query by path. Paths should not include the `[Workbook]` node.

```csharp
Query("myQueries\getDepartments")	//fetch the embedded query
	.Run() 							//run the query
	.Items 							//return the result rows (dynamic[])
```

### Range(...)
Gets the selected or specified range object.
```csharp
//get the range "a1:b2" and merge its cells into one.
Range("a1:b2").Merge();
```

### RunVBA(...)
Runs the specified VBA function or subroutine, passing in the specified parameters.
```csharp
//Assumes there's a function called "Add" that takes two numbers. 
//Important: the function should be defined in a module if it needs to return a result.
return Vba("Add", 1, 2);

//Assumes there's a subroutine "Abc" in Sheet1, taking one string parameter.
Vba("Sheet1.Abc", "Antonio");
```

### ShowDialog(...)
Shows the dialog, making sure to set the Excel window as its parent. Can be a custom form or e.g. an OpenFileDialog, OpenFolderDialog etc. Should be used for all dialogs, to ensure that the owner window is properly set.

```csharp
using System.Windows.Forms;

//Show an OpenFileDialog
var dialog = new OpenFileDialog();
ShowDialog(dialog);
MessageBox.Show($"You have chosen file {dialog.FileName}");

//Show a Form
ShowDialog(new Form());
```

### Workbook(...)
Returns the workbook that this script is associated with. For an embedded query this will be the workbook that contains the query. For non-embedded queries, it will be the currently active workbook. The main purpose of this is to ensure that embedded queries actually work with the correct workbook in automation scenarios.

!!! Note
	 `Workbook()` is a method rather than a property simply to avoid potential naming collisions with any table/range names present in the workbook.


## Extension methods

### Update(...)

```csharp
//signature
void Update<R>(this IEnumerable<R> rows, Action<R> action, [bool commit = true])
```

Updates a collection of rows using the specified action, and optionally commits changes to the Excel table.

```csharp
//example
people
	.Where(p => p.BirthDate == DateTime.Today)
	.Update(p => p.Name = "*" + p.Name)
```

### Delete(...)

```csharp
//signature
void Delete<R>(this IEnumerable<R> rows, [commit = true])
```

Deletes the collection rows from the table and optionally commits the change to the Excel table.
```csharp
//delete rows where name is null or empty
people.Where(p => string.IsNullOrEmpty(p.Name)).Delete();

//...or more simply
people.Delete(p => string.IsNullOrEmpty(p.Name));
```

### Format(...)
 
```csharp
//signature (for rows)
void Format<R>(this IEnumerable<R> rows, Action<Range> action)

//signature (for ranges)
void Format<Range>(this IEnumerable<Range> ranges, Action<Range> action)
```

Performs a bulk action (e.g. applies formatting) on a collection of rows or ranges. Optimizes communication with Excel in order to minimize the number of calls to Excel (much faster than applying formatting row-by-row).

```csharp
//Rows: change the interior color for all rows where BirthDate is today
people
	.Where(p => p.BirthDate == DateTime.Today)
	.Format(range => range.Interior.ColorIndex = 34);

//Cells: change the interior of cells in range "A1:B2:
Cells("a1:b2")
	.Format(range => range.Interior.ColorIndex = 33);
```


### IntoTable(...)
 
```csharp
//signature (for new table)
T IntoTable<T>(this T data, string tableName, string defaultAddress = null, Workbook workbook = null) where T : IEnumerable

//signature (for existing table)
T IntoTable<T, R>(this T data, TableDataAccessor<R> table) where R : RowDataAccessor where T : IEnumerable
```

Writes the collection of objects into the specified Excel table. If the table does not exist its name and location can be specified as strings. If if does exist, the table object itself should be passed. 

!!! Note
	When pushing data into an exsiting table, the current rows of the table are deleted and the new rows are inserted. However, only the column names that match the properties of the objects being inserted are populated. Any calculated columns will be left intact. The `[System.ComponentModel.DisplayNameAttribute]` can be used to override the mapping between columns and properties. 

```csharp
//Rows: change the interior color for all rows where BirthDate is today
people
	.Where(p => p.BirthDate == DateTime.Today)
	.Format(range => range.Interior.ColorIndex = 34);

//Cells: change the interior of cells in range "A1:B2:
Cells("a1:b2")
	.Format(range => range.Interior.ColorIndex = 33);
```



### IntoRange(...)
 
```csharp
//signature (for new table)
T IntoRange<T>(this T data, string address = null, bool includeHeader = false, Workbook workbook = null)
```

Writes the data into the range at the specified address. The data can be any object (does not need to be a collection). 

!!! Note
	The workbook parameter is optional and, when not specified, assumes the current workbook. In automation scenarios, especially if the query is triggered by a timer, `Workbook()` should be passed in order to make sure the correct workbook is used. Otherwise, a query running in a workbook that's in the background might write data into another workbook that's currently in the foreground.

```csharp
//Rows: change the interior color for all rows where BirthDate is today
people
	.Where(p => p.BirthDate == DateTime.Today)
	.Format(range => range.Interior.ColorIndex = 34);

//Cells: change the interior of cells in range "A1:B2:
Cells("a1:b2")
	.Format(range => range.Interior.ColorIndex = 33);
```
### IntoClipboard(...)

```csharp
//signature
T IntoClipboard<T>(this T data, bool includeHeader = true, string separator = null, bool useOAdataForDates = false)
```
Writes the data into the clipboard.

```csharp

people.IntoClipboard();

```