# Excel functions via SQL

In QueryStorm, you can define Excel functions that use SQL to return data from a database. 

Suppose we'd like to define a function that would return a list of people from a database, whose names contain a particular `searchTerm`.

We can achieve this by defining a new script that connects to the desired database and has the following code:

```sql
{function searchPeople(string searchTerm = "Joh")}

select
	*
from
	Person.Person p
where
	FirstName like '%' + {searchTerm} + '%'
```

We can run the query to examine if it returns the expected results. When running the query, the parameter's default value will be used. In this example, the query will return all persons whose name contains the string `"Joh"`.

To make this function available in Excel, we must:
1. **save** the script [*Ctrl+S*] so that it gets added to the project
2. **build** the project [*Ctrl+Shift+B*], to allow QueryStorm to load it

Click below to view a video demonstration:

[![Excel functions via SQL](../images/video.jpg)](https://youtu.be/rmya2vbUv18 "Defining Excel functions via SQL")

## Parameters

Parameters have names, types and default values. There are five different data types supported for parameters: int, float, datetime, string and bool.

The following are examples of valid parameter declarations:
- `int abc = 123`
- `float abc = 12.3`
- `datetime abc = "2020-08-31"` (ISO 8601 format)
- `string abc = "some text"`
- `bool abc = true`
- `var abc = "example text"` (auto-detect type)

To reference a variable in the body of the SQL query, you can either use preprocessor expressions e.g. `{abc}` or the normal parameter syntax for the database you are using, e.g. `@abc` (for SQL Server).

A parameter's default value is only used when the query is directly executed. When the function is called from Excel, the default value of the parameter is not used.  

## Return values

Functions can return a single value or an entire table as their result. 

If your machine (or the end user's machine) is running one of the newer (Office365) versions of Excel that support [dynamic arrays][1], the tabular result will automatically spill. 

If you are using an older version of Excel, you will need to use Ctrl+Alt+Enter (or `{=function()}` syntax) to allow the result to return a table of data (however, you'll need to know in advance the size of the output data). 

## Sharing functions

So once you've built the function, where can you use it? It depends on where you've created it. Functions can be defined inside a particular workbook or as a QueryStorm extension that's available in all workbooks. 

### Functions defined in the workbook

If the function is stored inside the workbook, anyone who has the workbook (and the QueryStorm runtime) will be able to use it. However, this function will only be usable in the workbook that contains it. If the function should be usable in any Excel file, you should define it in an "Extension" package (as described below).

QueryStorm supports building extension projects which can contain a set of user defined Excel functions that can be packaged together and published for use by other users.

To define the function in a QueryStorm extension rather than in the workbook, create a new project in "Code explorer" and then add a new script from the context menu of the new project.

For a video demonstration of this process, click below:

[![Excel functions via SQL (Extension package)](../images/video.jpg)](https://youtu.be/9mhYVjngI5w "Defining Excel functions via SQL (Extension package)")


[1]: https://support.microsoft.com/en-us/office/dynamic-array-formulas-in-non-dynamic-aware-excel-696e164e-306b-4282-ae9d-aa88f5502fa2 