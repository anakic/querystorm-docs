# Frequently asked questions

## General
- [What SQL dialect should I use, can I run complex SQL queries?](#user-content-what-sql-dialect-should-i-use-can-i-run-complex-sql-queries)
- [Will QueryStorm work well with huge workbooks?](#user-content-will-querystorm-work-well-with-huge-workbooks)
- [I don't see Oracle in the list of supported databases, can I connect to it?](#user-content-i-dont-see-oracle-in-the-list-of-supported-databases-can-i-connect-to-it)
- [Will this work with Excel for Mac?](#user-content-will-this-work-with-excel-for-mac)
- [What does embedding a query do to the workbook, can this workbook be used without QueryStorm?](#user-content-what-does-embedding-a-query-do-to-the-workbook-can-this-workbook-be-used-without-querystorm)
- [Can I roll back to a previous version temporarily, just in case the current version has an issue?](#user-content-can-i-roll-back-to-a-previous-version-temporarily-just-in-case-the-current-version-has-an-issue)
- [I need extra functionality from the VBA API, can you add it to the API?](#user-content-i-need-extra-functionality-from-the-vba-api-can-you-add-it-to-the-api)

## Installation
- [Do I need admin privileges to install QueryStorm?](#user-content-do-i-need-admin-privileges-to-install-querystorm)
- [I'm having installation issues, are there any easy steps I can try?](#user-content-im-having-installation-issues-are-there-any-easy-steps-i-can-try) 
  
## How-to
- [How do I update a database table with data from Excel?](#user-content-how-do-i-update-a-database-table-with-data-from-excel)
- [How do I use cell values as parameters in my query?](#user-content-how-do-i-use-cell-values-as-parameters-in-my-query)
- [How do I refresh a table (that was created from the results of a query)?](#user-content-how-do-i-refresh-a-table-that-was-created-from-the-results-of-a-query)
- [How do I update an Excel table with data from a database?](#user-content-how-do-i-update-an-excel-table-with-data-from-a-database)
- [How do I change row color using SQL (conditional formatting)?](#user-content-how-do-i-change-row-color-using-sql-conditional-formatting)
- [How do I create my own custom functions when I'm in *Live mode*?](#user-content-how-do-i-create-my-own-custom-functions-when-im-in-live-mode)
- [How do I call .NET functionality from SQL when I'm in *Live mode*?](#user-content-how-do-i-call-net-functionality-from-sql-when-im-in-live-mode)

## Pricing and licensing
- [Can I transfer a license from one machine to another?](#user-content-can-i-transfer-a-license-from-one-machine-to-another)
- [My workbook uses QueryStorm and I want to distribute it to my clients. Would they need to buy QueryStorm to use the workbook](#user-content-my-workbook-uses-querystorm-and-i-want-to-distribute-it-to-my-clients-would-they-need-to-buy-querystorm-to-use-the-workbook)
- [Is my license valid for future versions?](#user-content-is-my-license-valid-for-future-versions)

## Philosophical
- [Where does this fit in with PowerQuery and PowerPivot?](#user-content-where-does-this-fit-in-with-powerquery-and-powerpivot)
- [How does this compare to MS Query and VBA+ADO?](#user-content-how-does-this-compare-to-ms-query-and-vbaado)
- [Why is QueryStorm's VBA API asynchronous?](#user-content-why-is-querystorms-vba-api-asynchronous)

## Troubleshooting
- [Why isn't auto-complete offering any columns after `SELECT`?](#user-content-why-isnt-auto-complete-offering-any-columns-after-select) 
- [Why is my VBA function returning `Nothing` when called from SQL?](#user-content-why-is-my-vba-function-returning-nothing-when-called-from-sql)
- [Why can't I create additional indexes in *Live mode*?](#user-content-why-cant-i-create-additional-indexes-in-live-mode)
- [Why is the VBA callback I specified not being called after running a query?](#user-content-why-is-the-vba-callback-i-specified-not-being-called-after-running-a-query)

---

## General

### What SQL dialect should I use, can I run complex SQL queries?
Dialect depends on which engine you're using. Live mode uses the SQLite engine. In external mode you choose the engine, so if, for example,  you connect to Sql Server, you would use TSQL to write your queries. The queries can be as complex as you like.

### Will QueryStorm work well with huge workbooks?
If you are using the x64 version of Office it will work fine even with huge workbooks. In the x86 versions it might hit the 2GB memory limitation depending on what you're doing. In Live mode, columns are indexed implicitly ensuring response searches are snappy, but returning huge results sets will still take a while.

### I don't see Oracle in the list of supported databases, can I connect to it?
Well, sort of... Oracle does not support per-session temp tables, which are required for QueryStorm to copy workbook data to the database. This is why there is no Oracle specific engine. The fallback you can use is the ODBC engine, but it will not copy the excel tables, so you can use it to get data from Oracle but not into it.   

### Will this work with Excel for Mac?
Unfortunately, no. The *VSTO* runtime that QueryStorm relies on to talk to Excel doesn't exist for the Mac version of Office. 

### What does embedding a query do to the workbook, can this workbook be used without QueryStorm?
Embedding queries doesn't do anything funky with your workbook, it just writes a bit of text into a hidden area (xls and xlsx files have this area, and it's called CustomXMLParts). The workbook will show up perfectly fine even on machines that don't have QueryStorm installed. If there's no QueryStorm, then automation will not run, but the workbook will be perfectly fine and will work like any other Excel workbook.

### Can I roll back to a previous version temporarily, just in case the current version has an issue?
Yes, old versions can be downloaded from the [Versions](https://www.querystorm.com/versions.html) page. To install an old version, you will first need to uninstall the current one (Control Panel -> Add/Remove).   

### I need extra functionality from the VBA API, can you add it to the API?
Get in touch and suggest it, if it's reasonable, then absolutely. The API will grow and improve based on feedback.

## Installation
### Do I need admin privileges to install QueryStorm?
No, but it has two prerequisites which it will install if they are missing. These two prerequisites do require admin rights but are already present on most machines. 

The prerequisites are:
- .NET framework 4.5 (comes with Windows 8.0 or higher)
- Visual Studio Tools for Office Runtime (comes with Office 2013, 2016 and some versions of 2010)   

### I'm having installation issues, are there any easy steps I can try?
QueryStorm has two prerequisites as mentioned in the previous answer (.NET 4.5 and VSTO Runtime). If they are missing the installer will download and install them. If the installation is failing, it might be due to errors installing .NET4.5 or vsto_runtime. Installing them manually (both can be downloaded from Microsoft for free) will most likely resolve any installation issues you might have.  


## How-to

### How do I update a database table with data from Excel?
As soon as you connect to the database, the selected tables will be in the destination database as temp tables. Just move the data from there using SQL (`select...into...` or `merge` or whatever is appropriate). Pretty simple. 

### How do I use cell values as parameters in my query?
Use curly brackets, the preprocessor will replace them with values from cells. You can use addresses and named ranges. Example:
```sql
select * from person where year(birthdate) = {Sheet1!b1} 
```

### How do I refresh a table (that was created from the results of a query)?
Short answer: From the table's context menu-> Related Queries -> Run. 

Long answer: If you created the table using *table writer*, and `SetupRefresh` was set to `Manual` or `Automatic`, then the query will already be embedded and associated with the table, and will appear in the context menu of the table (under Related queries). You can run it from there. 

You can also manually associate any embedded query with a table by making sure the name of the query begins with '*[nameOfTable].*'. The query will then appear in the context menu of the table. 

You can also add other triggers for executing the query automatically (like an ActiveX button, or a timer). For info on this take a look at the *Automation *section.


### How do I update an Excel table with data from a database?
Cookbook for this scenario:

1. Connect to the database, make sure to select the table so it gets copied as a temp table to the destination database
2. Use data from the database to update the temp table 
3. Select everything from the temp table and use @-directive to redirect the result back into the Excel table 

Example:
```sql
/* some code that manipulates the data in the #Person table */
update #Person 
set name='Percival' where Age=32 and Department='King Arthur''s legendary Knights of the Round Table'

/*select the modified data but redirect the results into the workbook table*/
@Table TableName="Person"
SELECT * FROM #Person
```

Alternatively, don't manipulate the temp table, just write a query that gives you back the results and redirect the results using an @-directive into the Excel table. Just make sure that the columns names of the workbook table and the columns names in the results match (only the columns that you want to update need to match).

More on this in *Writing results* and *Preprocessor* sections.    

### How do I change row color using SQL (conditional formatting)?
In Live mode, there are several functions that can manipulate the font color and background color of an Excel range (documented in *Guide/Live mode functions*). The functions need to know the row ranges in Excel, which is why all tables in QueryStorm have a hidden column called `__address` which specifies the original address of each row in Excel. We use this to manipulate the formatting of rows in an Excel table. 

Here's an example:
```sql
SELECT
	CASE JobTitle
		WHEN 'Chief Executive Officer' THEN SetBackgroundColorRGB(__address, 255, 0, 0)
		WHEN 'Vice President of Engineering' THEN SetBackgroundColorRGB(__address, 255, 255, 0)
		WHEN 'Research and Development Manager' THEN SetBackgroundColorRGB(__address, 0, 255, 255)
		ELSE ClearBackgroundColor(__address)
	END
FROM
	Employee
```   

### How do I create my own custom functions when I'm in *Live mode*?
There are two options for implementing user defined functions:
- implement the function in VBA and call it using the `vba(funcOrSub, arg1..argN)` function
- write it as a C#-like expression and run it using the `eval(expr, arg1..argN)` function (take a look at the next question) 

When calling a VBA function, make sure that the function is in a module (**not in the workbook or a sheet**), otherwise the return value will be `Nothing`. 

More details on both `vba` and `eval` functions in section *Functions in live mode*.


### How do I call .NET functionality from SQL when I'm in Live mode?
You can execute C#-like expressions in a query by using the `Eval(expr, arg1..argN)` function that's available in *Live mode*. The expression is basically a lambda expression (a function defined inline). The engine will run the expression for each row in the results as if it was a SQL function. The expression looks a lot like C#, you can use static and instance members of .NET classes, instantiate objects and use operators, but the expressions can't contain lambda expressions themselves, and they can contain parameter placeholders. Here's an example:

```sql
select *, eval('DateTime.Parse($0).DayOfWeek', dateofbirth) from Players  
```

Obviously this could have been done using the `DayOfWeek` function that's available in Live mode, but this is just for illustration. Also note that **SQLite does not have a DateTime type, dates are treated as strings** so they need to be parsed inside `Eval` to get the `DateTime` object. 


## Pricing and licensing

### Can I transfer a license from one machine to another?
Yes, absolutely. The licensing UI allows moving a license, it's a simple process that you can do yourself and usually only takes about a minute. If you have any trouble with this, please get in touch.

### My workbook uses QueryStorm and I want to distribute it to my clients. Would they need to buy QueryStorm to use the workbook?
If the embedded queries use external database connections, then yes. That said, this feature will either be free in the long term or there will be an affordable runtime-only license and a *headless* version of the plugin for distributing to clients.

### Is my license valid for future versions?
Yes, you will be able to use all future versions up to version 2. FYI there are currently no plans for switching to version 2, all upgrades in the foreseeable future will still be versions 1.[x.y.z].  


## Philosophical

### Where does this fit in with PowerQuery and PowerPivot?
PowerPivot and PowerQuery are BI tools, and even though BI operations and SQL operations do somewhat overlap, they are appropriate for different kinds of tasks. 
QueryStorm actually has more in common with SQL Server Management Studio than with PowerPivot. 
It's main strengths is allowing the use of SQL for cleaning, querying and manipulating data. 
PowerPivot on the other hand specializes in aggregating data, slicing and dicing, but for precise querying, cleaning and molding data, SQL is more appropriate and much more expressive. 
QueryStorm offers some unique possibilities you can benefit from regardless of whether or not you're using PowerQuery and PowerPivot.
  
### How does this compare to MS Query and VBA+ADO?
Running SQL queries against sheets and named ranges has been possible with MS Query and ADO for a long time using JET/ACE engines. 
Both MS Query and VBA/ADO are used somewhat rarely though, and the reason is that it's not easy, it's somewhat limited and involves a lot of ceremony. 
QueryStorm aims to provide an environment where SQL is convenient and less constrained so it can be used easily, effectively and often.

### Why is QueryStorm's VBA API asynchronous?
Query execution in QueryStorm is completely asynchronous in order to allow the user to interact with Excel while long running queries are executing. If the query execution was synchronous the user wouldn't even be able to cancel query execution because the UI would not be responding while the query was running. The VBA API uses the existing execution runtime which just doesn't have any synchronous functionality. That said, if this turns out to be an unreasonable limitation, synchronous API methods might get added in the future.    

## Troubleshooting

### Why isn't auto-complete offering any columns after `SELECT`?
When entering columns in the select list, only columns belonging to objects in the FROM clause will be offered, which is why entering the FROM clause first is required for column completion in the select list. It's a bit awkward that in SQL, SELECT comes before FROM, as it makes auto-complete much less helpful (without knowing the FROM clause, auto-complete cannot offer meaningful columns in the select list).

### Why is my VBA function returning `Nothing` when called from SQL?
Functions need to be inside modules to return values. If they are in the sheet or the workbook, they will not be able to return a value to SQL and will return `Nothing`.

### Why can't I create additional indexes in *Live mode*?
SQLite doesn't allow creating indexes for virtual tables (the mechanism used to enable it to work with Excel tables). There is a way around this though. If you need custom indexes, use the SQLite engine as an external db (Connect custom -> SQLite), when connecting use the following connection string to use it in-memory:

    Data Source=:memory:

This way, the tables are normal database temp tables (not virtual tables), and you can create any indexes you need on them. 

Modifying the data inside them will not change the data in Excel tables, but if you need to update excel tables, you can do it using the table writer (`@Table`), as you normally would in external mode.  

### Why is the VBA callback I specified not being called (after running a query)?
The likeliest reason is that you forgot to qualify it. If it is not in a module, it needs to be qualified, e.g. `query.Run("Sheet1.CallBackName", "Sheet1.ErrorCallbackName")`. If that's not the cause, then it's a bug :)