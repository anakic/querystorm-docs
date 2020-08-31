# Excel functions via SQL

You can use SQL to define Excel functions. A typical use case for this would be to pull in data from a database. 

Suppose we want to define a function that will accept a string parameter called `searchTerm` and return a list of people from a SQL Server database whose names contain that string.

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

We can run the query to examine if it returns the expected results. When running the query, the parameter's default value will be used, and the query will return all persons whose name contains the string `"Joh"`.

To make this function available for use from Excel, we must first:
1. save it [*Ctrl+S*]
2. build the project [*Ctrl+Shift+B*].

We need to save the file to ensure that it's saved into the project. We need to build the project in order to allow QueryStorm to load functions from it.


Click below to view a video of the process:
[![Everything Is AWESOME](https://encrypted-tbn0.gstatic.com/images?q=tbn%3AANd9GcTY-5zRYwmgJFGuWvZxc8kSKnSksrbTB5183Q&usqp=CAU)](https://youtu.be/rmya2vbUv18 "Defining Excel functions via SQL")


## Parameters

Parameters have types and default values. They can be declared in the following ways:
- `int abc = 123`
- `float abc = 12.3`
- `datetime abc = "2020-08-31"` (ISO 8601 format)
- `string abc = "some text"`
- `bool abc = true`
- `var abc = "example text"` (auto-detect type)

A parameter's default value is only used when the query is directly executed. When the function is called from Excel, the default value of the parameter is not used.  

## Return values

Functions can return a single result or an entire table. 

If your machine (or the end user's machine) is running one of the newer (Office365) versions of Excel that support dynamic arrays[^1] the results of functions that return a tabular result will automatically spill.


[^1]: See dynamic arrays 
