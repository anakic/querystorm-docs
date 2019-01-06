# Connecting to external databases

In general, tech people are effective at dealing with data in databases, but are not as effective when working with data in spreadsheets. Still, a lot of data gets passed around and consumed via spreadsheets, so the ability to move data easily between spreadsheets and databases can be extremely useful. 

QueryStorm makes this very easy. As soon as you connect to an external database, selected workbook tables are visible to the database as temp tables. From there, the data can be queried or imported into permanent tables. In the other direction, query results can be returned into the workbook and used to create new tables or update existing ones.

QueryStorm supports connecting to SQL Server, PostgreSQL, MySql, Access, SQLite, DB2 and Redshift with limited support for Oracle and ODBC connections.

!!! Note
	Oracle and ODBC connections are limited to reading data - they do not allow importing Excel tables as temp tables. The reason is that Oracle does not support session-level temp tables, while ODBC does not perscribe the syntax for creating temp tables (syntax depends on the destination database).

In this article you'll learn how to:

- Connect to external databases.
- Move data between Excel and databases in either direction.
- Create interactive workbooks that load data from databases automatically.

For a video tutorial about working with external databases, click [here](https://www.querystorm.com/tutorials.html#217163949 "Connecting to external databases video tutorial").

## Querying

To connect to an external DB engine, click the **Connect with** button and configure your connection:

![Connecting to external SQL Server](../images/connect_custom_dialog.png "Connecting to external SQL Server")

Selected workbook tables will be copied to the destination database **as temp tables**. Once connected, we can query the Excel data alongside existing database data, using all the features and processing power of the database server. 

![Connected external](../images/connected_external.png "Connected to external SQL Server")

In this example, we're taking the *salaries* table from Excel and making it available to the database as a temp table. 

We can query the workbook tables alongside permanent database tables. For example, the following query will search for people who are present both in the *salaries* table in Excel and in the *persons* table in the database:

``` SQL
SELECT *
FROM #salaries s
WHERE EXISTS (
		SELECT *
		FROM Person.Person p
		WHERE s.EmployeeName = p.FirstName + ' ' + p.LastName)
```

## Getting data into Excel

### Writing results into a new Excel table

Query results can be written into Excel ranges and tables using the [preprocessor](../misc/preprocessor):

```sql
@Table TableName="dpt" DefaultAddress="Sheet1!B2"
SELECT * FROM HumanResources.Department
```   
The results of the above select statement will be written into an Excel table called *dpt*. If the table does not yet exist, it will be created at the location `B2` in *Sheet1*. If the table exists, it will be overwritten by the data in the results, but any calculated columns will be left intact.

### Deleting rows from Excel tables

The database engine only has access to a copy of the data in the Excel tables. It can't manipulate the Excel tables directly via SQL. However, we can use the preprocessor to overwrite Excel tables with updated data from the database.

Suppose we want to delete some rows from an Excel table. We can delete the rows in the temp table, and then push the data from the temp table into Excel:

```sql
delete from #myTable where id > 10

@Table TableName="myTable"
SELECT * FROM #myTable
```

Alternatively, we could achieve the same result but pushing just the desired subset of data, without modifying the temp table:

```sql
@Table TableName="myTable"
SELECT * FROM #myTable where id <= 10
```  

### Updating data in Excel tables

Updating data works similarly to deleting. We can modify the data in the temp table and then output the updated data into the Excel table:
```sql
UPDATE #myTable SET Id = Id + 10

@Table TableName="myTable"
SELECT * FROM #myTable
```  

Alternatively, we could achieve the same result without modifying the temp table by returning new data for one or more columns:  
```sql
@Table TableName="myTable"
SELECT 
	Id + 10 as Id -- must make sure the column is still named "Id"  
FROM 
	#myTable
```
!!! Note
	Column names in the results should match the column names in Excel in order to update them. Unmatched columns are left intact.

## Importing Excel data into permanent tables
Excel data is available to the database in the form of temp tables. From there, it's quite easy to move it into permanent database tables. Here are some examples.

To create a new database table from the imported Excel table:
```sql
-- [SQL Server syntax]
SELECT *
INTO dbo.my_new_permanent_table
FROM #my_Excel_table
```

To insert data into an existing table:
```sql
insert into dbo.my_existing_permanent_table(a,b,c)
select a,b,c from #my_Excel_table 
```
!!! Tip
	Since the temp table contains extra columns (`__address` and `__row`) we can't use `*` in the select list, but we can easily expand the `*` in the editor by invoking autocomplete on it and choosing the *[expand]* option. 

For more complicated scenarios, multiple statements might be appropriate (e.g. when distributing the data into multiple database tables).


## Special columns

As with the built-in SQLite engine, each table gets two extra columns `__address` and `__row`. These columns exist in the temp tables that hold the data imported from Excel. Databases other than SQLite don't have a concept of hidden columns so for these databases, these are just regular columns.

**The `__address` column** contains the original address of each row in Excel. This makes it easy to find particular rows in Excel since double-clicking an address in the results grid will highlight the row in Excel.

**The `__row` column** contains the index of a row in the Excel table. Databases don't guarantee the order of rows in the results, so this column is added to enable preserving the original ordering of rows. It can also be useful for certain types of queries such as self-joins without duplicates (e.g. join each row with all rows below it).


## Refreshing data
QueryStorm supports automation which enables building interactive workbooks that run scripts and queries in the background as the user interacts with the workbook.

For example, the user changes a date in a cell which triggers a query to run in the background which fetches a list of transactions for that date from the database. This data then gets pushed into an Excel table which, in turn, updates a graph. 

You can read more about automating queries in the [automation section](../automation/embedding_code) of this documentation. 


## Parameterizing queries
SQL queries in QueryStorm can use cells as parameters. This is especially useful in automation scenarios.

Here's how to use a value from cell `A1` as a parameter in your query:

``` sql
select * from People where Id = {Sheet1!A1}
```

You can read more about parameterizing queries [here](../misc/preprocessor/#parameterizing-queries). 

## Built-in engine vs External engines
### Advantages of the built-in SQLite engine
- No need for having an external database engine at hand
- Can use SQL `insert`/`update`/`delete` commands to modify data in workbook tables
- Can run functions that modify row formatting from SQL (e.g. `SetBackgroundColor`)
- Can call C# and VBA functions from SQL

### Advantages of using external databases
- Can combine data from workbook tables with data from database tables in the same query
- Can use server-specific functionality when processing data (e.g. window functions) 
- Can import data from a database into the workbook (e.g. for reporting)
- Can import data from Excel into the database