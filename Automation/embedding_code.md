# Embedding code

One-off queries and scripts are great for cleaning and querying data, but if we want to give our workbooks interactive behavior and make the data they contain refreshable, we need to make our code part of the workbook.

## How to embed a query/script?

To embed a SQL query or C# script into the workbook, right click in the code editor and click *Embed*:

![Embed query](../images/embed_command.png "Embed query")

Alternatively, we can click the *New query* button in the *Workbook queries* pane. 

![New embedded query](../images/embed_new.png "New embedded query")

!!! Naming
	If you have a table called `Table1` and your embedded query is called `Table1.TestQuery`, the query will show up in the context menu of Table1 in the "Related queries" submenu. You can rename your query by pressing F2 or clicking the *Rename* button.


## Engine configuration
Aside from code, each embedded query/script also has information on the connection that should be used when executing it. When you create an embedded query, it automatically picks up the configuration of the current connection. 

To view and change the connection configuration, select the query in the *Workbook queries* pane and click the *Connection* button.

![Connection edit button](../images/connection_edit_button.png "Connection edit button")

This will open a dialog which will let you configure the setting for the connection. 

![Connection edit dialog](../images/connection_edit_dialog.png "Connection edit dialog")

In the example above, the connection configuration includes the following information:

1. Type of connection: `MsSql` (SQL Server)
1. the Excel tables to import as temp tables
1. the connection string: `Server=.\sqlexpress; Database=AdventureWorks2014; Trusted_Connection=True;`

You can edit the connection configuration manually, or use one of the two buttons to use the current connection or choose another configuration via the *Connect custom* dialog.

!!! Warning
	Make sure to **NOT include any Excel tables you don't use in your query**, because copying data from Excel tables can slow things down when connecting, especially if the table is large.

## When to embed?

The point of embedding a query is to enable it to be executed again at a later time. A query can be included as an example that someone else can view and run manually on their own machine. 

More often, though, queries are embedded so they can be automated in order to add interactive behavior to the workbook. In that case, the end user of the workbook need not be technical - they just interact with the workbook and the code runs in the background (e.g. refreshes data, performs actions...). 

## Running embedded queries

An embedded query or script can be executed in several ways:

- If it's associated with a table (via naming convention), we can run it from the context menu of the table.
- We can run it manually by clicking the *Run embedded query* button in the *Workbook queries* pane
- It can be set to run automatically when triggered, as described in the [next chapter](jobs.md). 
