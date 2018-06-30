# Connecting to external databases

Excel is more useful if we can easily get data in and out of it, primarily in relation to databases. 

QueryStorm make this much easier than other tools. It allows moving live data from Excel into the database, as well as fetching results back into Excel with arbitrary queries (for which you get good IDE support).

In this article you'll learn how to:
- connect to external database servers and use them to process your Excel data
- move data between Excel and a database in either direction
- create interactive workbooks that load data from databases automatically

## Querying

To connect to an external db engine, click the **Connect with** button and configure your connection:

![Data in a sheet](https://www.querystorm.com/Images/docs/connectcustom.png?v=1)

Selected workbook tables will be copied to the destination database **as temp tables**. Once connected, we can query the Excel data alongside existing database data, using all the features and processing power of the database server. 

To demonstrate querying Excel data alongside database data, let's search for people who are present both in the `salaries` table in Excel and in the `persons` table in the database:

``` SQL
SELECT *
FROM #salaries s
WHERE EXISTS (
		SELECT *
		FROM Person.Person p
		WHERE s.EmployeeName = p.FirstName + ' ' + p.LastName)
```

See this [tutorial](https://www.querystorm.com/tutorials.html#217163949 "Connecting to external databases video tutorial") for a video demonstration.

## Getting data into Excel

### Writing results into a new Excel table

Query results can be written into Excel ranges and tables in [several ways](https://www.querystorm.com/docs/Writing_results). In this example, I'll show how to do it using the preprocessor:

```sql
@Table TableName="dpt" DefaultAddress="Sheet1!B2"
SELECT * FROM HumanResources.Department
```   
The query results will be written to an Excel table called *dpt*. If the table does not yet exist, it will be created at the location `B2` in *Sheet1*.

If the table exists, it will be overwritten by the data in the results, but any calculated columns will be left intact.

Since the database engine only has access to a snapshot of data from Excel tables, it can't manipulate the data in the Excel tables directly. However, using the preprocessor we can output updated data into Excel tables and, in effect, update them. 

### Deleting rows from Excel tables

So let's say we want to delete some rows from an Excel table. We could do it in one of two ways.


**Option A:** return a subset of rows and output the results into the Excel table:

```sql
@Table TableName="myTable"
SELECT * FROM #myTable where id > 10
```  

**Option B:** delete the rows in the temp table, and then output its contents into the Excel table:

```sql
delete from #myTable where id > 10

@Table TableName="myTable"
SELECT * FROM #myTable
```

### Updating data in Excel tables

Updating data works similarly to deleting... 

**Option A:** return new data for one or more columns, and output the result into Excel.  
```sql
@Table TableName="myTable"
SELECT Id + 10 as Id FROM #myTable
```
> Note #1: I'm keeping the column name the same via alias (`as Id`).
> Note #2: I don't need to include the columns I'm not updating, they'll be left intact.

**Option B:** modify the data in the temp table and then output its contents into Excel
```sql
UPDATE #myTable SET Id = Id + 10

@Table TableName="myTable"
SELECT * FROM #myTable
```  

## Importing Excel data into permanent tables
As mentioned, Excel data is made available in the database in the form of temp tables. From there, it's quite easy to move it into permanent database tables. Here are some examples.

To create a new database table from the imported Excel table:
```sql
-- [SQL Server syntax]
SELECT *
INTO dbo.my_new_permanent_table
FROM #my_Excel_table
```

To insert data into an existing table:
```sql
insert into dbo.my_new_permanent_table(a,b,c)
select a,b,c from #my_Excel_table 
```
> Note: since the imported Excel table contains extra columns (`__address` and `__row`) we can't use * in the select list, but we can easily expand the * in the editor by invoking autocomplete on it and choosing the *[expand]* option. 


For more complicated scenarios, multiple statements might be appropriate (e.g. when distributing the data into multiple database tables).


## Special columns

As with the built in SQLite engine, each table gets two extra columns `__address` and `__row`. These columns exist in the temp tables that hold the data imported from Excel. 

**The `__address` column** contains the original address of each row in Excel. This makes it easy to find particular rows in Excel, since double-clicking an address in the results grid will highlight the row in Excel.

**The `__row` column** contains the index of a row in the Excel table. Databases don't guarantee the order of rows in the results, so this column is added to enable preserving the original ordering of rows. It can also be useful for certain types of queries such as self join without duplicates (e.g. join each row with all rows below it).


## Refreshing data
QueryStorm supports automation, so queries can be embedded and executed automatically when specific events occur. The point of this is to enable building interactive workbooks that run scripts and queries in the background as the user interacts with the workbook.

This enables building application/dashboard style workbooks.

The [automation section](http://docs.querystorm.com/automation) describes how to embed queries and set up jobs and triggers. 


## Parameterizing queries
SQL queries in QueryStorm can use cells as parameters. This is especially useful in automation scenarios.

Here's how to use a value from cell as a parameter in a query:

``` sql
select * from People where Id = {Sheet1!A1}
```

You can read more about parameterizing queries [here](http://docs.querystorm.com/Preprocessor#Parameterizingqueries). 

## Built-in vs External engines
### Advantages of the built-in SQLite engine
- No need for having an external database engine at hand
- Can use SQL insert/update/delete commands to modify data in workbook tables
- Can run functions that modify row formatting from SQL (e.g. `SetBackgroundColor`)
- Can call VBA functions from SQL

### Advantages of using external databases
- Can combine data from workbook tables with data from database tables in the same query
- Can use server specific functionality when processing data (e.g. window functions) 
- Can import data from a database into the workbook
- Can export data from Excel into a database  