# Welcome to QueryStorm

This document is a 2-minute introduction to help you quickly get started with QueryStorm.

## Tables, not sheets
QueryStorm works with [Excel tables](https://support.office.com/en-us/article/overview-of-excel-tables-7ab0bb7d-3a9e-4b56-a3c9-6c94334e492c "Excel tables") (not sheets). This is an important distinction, so I mention it a few more times throughout this documentation.

![Tables](images/tables.png)

!!! Note
	Press Ctr+T on a cell or a block of cells to convert it to a table.

## Querying with SQL
Once you have the data in tables, you can start querying by clicking **Connect (SQLite)** in the ribbon. This will pop up the QueryStorm IDE and open a SQLite connection that *sees* Excel tables as database tables. 

![Querying with SQL](../images/Untitled-Project.gif)

The database engine is [SQLite](https://www.sqlite.org) and you can run any valid SQLite query on your data, including `UPDATE`, `INSERT` and `DELETE` statements. You can even [change row formatting](../sql/formatting) via SQL. 

Read more about the SQL support [here](../sql/querying).

## Querying with C# #
Aside from SQL, you can also query data via C#. Click the **Connect (C#)** button in the ribbon to open the IDE and start an interactive C# session. Excel tables will show up as strongly typed collections that you can query and modify. Named ranges will also show up as strongly typed variables which you can use to get/set the values of the named ranges. 

![Querying with C#](images/csharpintro.gif)

Read more about QueryStorm's C# support [here](../csharp/querying).

## Connecting to external databases

You can connect to external databases and move data in either direction. This makes importing, exporting and combining database and spreadsheet data much easier. 

To connect to an external database, click **Connect custom**. A dialog will appear allowing you to choose (or define) a database connection:

![Defining a connection](images/connectivity.png)

While setting up the connection, you select which Excel tables need to be visible to the database. These tables are copied to the database as temp tables as soon as the connection is established.

![Connected to DB](https://i.imgur.com/3Fqom4n.png)

You can now query the data from Excel tables alongside permanent database tables, using all the resources of the database server. Query results can easily be returned into Excel and used to populate tables or ranges.

Read more about working with external databases [here](../external/databases).

## Try it!
Congratulations, you now know the basics of QueryStorm. There's a lot more interesting stuff in the rest the documentation and some [video tutorials](https://www.querystorm.com/tutorials.html) as well, so I encourage you to take a look.   

