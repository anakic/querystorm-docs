# Writing and refreshing results
There are several ways of writing query results into Excel. The simplest one would be to just copy and paste them from the results grid, but if you need to do it repeatedly, this isn't a very good option. 

Also, if you're trying to update an existing table with a result-set that has a different number of columns (e.g. the table has extra calculated columns), copy-pasting can get messy quite quickly.

QueryStorm provides mechanisms for writing results into tables, ranges and the clipboard. All three can be scripted(automated) using the [preprocessor](https://www.querystorm.com/docs/Preprocessor "Preprocessor").  

To use the result writers, click the "*Write results*" button in the ribbon. It will be enabled only if the current query has any results. 

![Data in a sheet](https://www.querystorm.com/Images/Docs/writeresultsbtn.png)

## Table writer
This is the most complex and important writer out of the three. It can create new tables, update or recreate existing ones, as well as set up *refreshability*.

![Table writer](https://i.imgur.com/CmKDdt8.png)

All of the settings in the dialog have tool-tips that describe them (the blue icon to the right), but the most important ones are also described here in a bit more detail:

- **Table name**: can point to an existing table to update or recreate it, or a new name if creating a new table.
- **If exists**: specifies what happens in case *TableName* points to an existing table. 
	- `Update`: overwrites the data in the table but leaves calculated columns (and other columns not present in the resultset) intact.
	- `Recreate`: the entire table is deleted and recreated while keeping any formatting information intact (table style, column formats). 
- **Default address**: the location where the table will be created if `Table name` does not point to an existing table.
- **Embed query**: if checked, also embeds the query to enable refreshing the data.
- **Auto create triggers**: detects any Excel cells or tables in the query, and sets them as triggers for executing the embedded query.  

### Scripting the table writer
To script the table writer, use the `@Table` [preprocessor directive](https://www.querystorm.com/docs/Preprocessor#Redirectingresults "Redirecting results into table via preprocessor") like so:
``` sql
@Table TableName="People"
select * from humanresources.person
```
This redirects the results of the query into an Excel table called *"People"*. 
The properties of the @Table directive correspond to the properties in the dialog, and there is also autocomplete support for them.

### Refreshing the table
If `Embed query` was checked, aside from writing the data, the query is also embedded into the workbook. The name of the query will have the following naming convention `[name of table].[name of query]`, for example `movies.Refresh`.

This naming convention ensures that the query appears as a related query in the context menu of the table. 

![Related queries context menu](https://i.imgur.com/KzKP8T0.png)

From here, we can re-run the query manually, but more importantly, once the query is embedded, we can set up automation, so that it gets executed automatically when appropriate.

## Range writer
This writer is useful for writing a scalar result from an external database into a cell. 

Since it's primarily used for writing a single value, it's usually only useful when scripted, in automation scenarios. 

For example:
``` sql
@Range Address="Sheet1!F14"
select count(*) from humanresources.person
```

## Clipboard writer
The clipboard writer is useful mainly to prepare text to paste into another application or a file.