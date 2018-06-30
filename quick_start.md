# Welcome to QueryStorm

This document is a 2-minute introduction to the basics of QueryStorm to help you get started quickly.

## Tables, not sheets
QueryStorm works with Excel tables [[not to be confused with sheets](https://support.office.com/en-us/article/overview-of-excel-tables-7ab0bb7d-3a9e-4b56-a3c9-6c94334e492c "Excel tables")].

![Tables](https://www.querystorm.com/downloads/images/tables.png)

!!! Note
	You can convert a range of cells into a table by selecting a cell in the block of cells that contains your data and pressing Ctrl+T, or clicking *Insert->Table* in the Excel ribbon. You can rename the table by selecting a cell inside it, and going into the Design tab in the ribbon.

## Querying with SQL
Once you have your data in tables, you can start querying by clicking **Connect (SQLite)** in the ribbon. Excel tables will be displayed in the object explorer and you can work with them as if they were database tables.  

![Querying with SQL](https://www.querystorm.com/images/Untitled-Project.gif)

The database engine is [SQLite](https://www.sqlite.org). You can run any valid SQLite query on your data. You can even [change row formatting](../sql/formatting) via SQL.

You can read more about SQL support [here](../sql/querying).

## Querying with C# #
You can connect via C# using the **Connect (C#)** button in the ribbon. Excel tables will show up as strongly typed collections that you can query and modify. Named ranges will also show up as variables which provide strongly typed access for getting/setting the value of the named range.

![Querying with C#](https://www.querystorm.com/downloads/images/csharpintro.gif)

You can read more about QueryStorm's C# support [here](../csharp/querying).

## Connecting to external databases

You can connect to external databases and move data in either direction. This makes importing, exporting and combining database and Excel data trivial. 

To connect to an external database, click **Connect custom** and define (or choose) a database connection:

![Defining a connection](https://www.querystorm.com/images/images/connectivity.png)

While setting up the connection, you can select which Excel tables to make available to the database (as temp tables). 

![Connected to DB](https://i.imgur.com/3Fqom4n.png)

Once connected, you can query the Excel tables alongside permanent database tables. You can move the Excel data from the temp tables into permanent tables if you need to import. Query results can also easily be returned back into Excel and used to populate tables or ranges.

Read more about working with external databases [here](../external/databases).

## Try it!
Now that you know the basics, you're ready to get querying. You can read the rest of the docs for more info on all of this. You like video tutorials, check them out [here](https://www.querystorm.com/tutorials.html).   

