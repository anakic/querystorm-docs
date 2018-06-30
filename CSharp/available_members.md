# APIs available to scripts

This document outlines the members and APIs that are available to user scripts.

## Excel tables
Each table is represented by a object derived from `TableDataAccessor<T>`. The following is a list of relevant members each table has:

### Range
Get the Range in Excel that contains the table.
```csharp
table.Range.Interior.ColorIndex = 0;
```

### ListObject
Get the ListObject that the table object refers to.
```csharp
table.ListObject.ShowTotals = true;
```

### SelectedRows
Returns a list of rows that are part of the current selection in Excel.
```csharp
table.SelectedRows.FirstOrDefault()?.Id
```

### SaveChanges()
Commits changes (updates, inserts and deletes) into the Excel table.
```csharp
table.First().Name ="abc"; 
table.SaveChanges();
```

### Insert(...args...)
Inserts a row with the specified values into the table.

```csharp
table.Insert(10, "abc");
```

### InsertIntoCache(...args...)
Inserts a row with the specified values into the table cache (does not commit to Excel). Useful for bulk inserting rows (saving changes to Excel is slow).

```csharp 
table.InsertIntoCache(10, "abc");
table.InsertIntoCache(11, "def");
table.InsertIntoCache(11, "ghi");
table.SaveChanges();
```

## Named ranges
Each named range is represented as a global get/set property that provides strongly typed access to the value of the named range. 

Suppose we have a named range *abc* that contains a numeric value.

Reading the value from the named range:
```csharp
return abc + 5;
```

Setting the value in the named range:
```csharp
abc = 123;
``` 

## Global methods
The following global methods are available to C# scripts:

### Application()
Returns the current Excel application.

```csharp
Application().Workbooks.Open(@"c:\workbook.xls");
```

!!! Note
	This is a method rather than a property simply in order to avoid potential naming collisions with table/range names.

### Cells([address])
Gets a collection of cells in the current selection or specified range.
```csharp
Cells().Format(r=>r.Interior.ColorIndex = 34);
Cells("a1:b2").Format(r=>r.Interior.ColorIndex = 27);
```

### Csv<T>(path, [delimiter], [culture])
Opens the specified CSV files and returns a list of rows, mapped to the specified class by header.

```csharp
//Get rows in the people.csv file, mapped to the Person class.
IEnumerable<Person> people1 = Csv<Person>(@"c:\people.csv");

//...also specify delimiter
IEnumerable<Person> people2 = Csv<Person>(@"c:\people2.csv", "\t");

//...also specify datetime/number formats
IEnumerable<Person> people3 = Csv<Person>(@"c:\people2.csv", "\t", System.Globalization.CultureInfo.GetCultureInfo("pl"));
```

### Log(value, [depth])
Writes the specified value into the QueryStorm messages log. Drills down through properties to the specified depth.

```csharp
Log(person, 1);
```

### Query(path)
Gets an embedded query by path. Paths should not include the `[Workbook]` node.

```csharp
Query("myQueries\getDepartments").Run().Items //returns the results of executing the "myQueries\getDepartments" query
```

### Range([address])
Gets the selected or specified range object.
```csharp
Range("a1:b2").Merge();
```

### RunVBA(name, parameters)
Runs the specified VBA function or subroutine, passing in the specified parameters.
```csharp
//Assumes there's a function called "Add" that takes two numbers defined in a referenced module.
return Vba("Add", 1, 2);

//Assumes there's a subroutine "Abc" in Sheet1, taking one string parameter.
Vba("Sheet1.Abc", "Antonio");
```

### ShowDialog(form)
Shows the dialog, ensuring that the Excel window is the parent. Can be a custom form or e.g. an OpenFileDialog, OpenFolderDialog etc. Should be used for all dialogs, otherwise the owner window of the dialog will not be properly set.

```csharp
using System.Windows.Forms;

//Open file dialog
var dialog = new OpenFileDialog();
ShowDialog(dialog);
MessageBox.Show($"You have chosen file {dialog.FileName}");

//Form
ShowDialog(new Form());
```

### Workbook()
Returns the workbook that this script is associated with. 

!!! Note
	This is a method rather than a property simply in order to avoid potential naming collisions with table/range names.


## Extension methods

| Type | Method | Description |
|---|---|---|
| `IEnumerable<TRow>` | `Update(action, commit)` | Updates a collection of rows using the specified action, and optionally commits changes to Excel table. |
| `IEnumerable<TRow>` | `Delete(commit)` | Updates a collection of rows using the specified action, and optionally commits changes to Excel table.|
| `IEnumerable<TRow>` | `Format(rangeAction)` | Apply formatting (in bulk) to the collection of rows.|
| `IEnumerable<TRow>` | `ToAddressGroup()` | Gets an `AddressGroup` object that represents a group of ranges. Bulk operations (e.g. setting background color) can be done much faster on `AddressGroup` compared to working with individual rows.|
| `IEnumerable<T>` | `Do()` | Similar to `ForEach`, but executes on the main thread (improved performance when interacting with Excel API) and automatically commits any changes to the source table if `T` is a row type.|