# Cross-database queries

As mentioned before, the SQLite engine in QueryStorm can access data from the external world via [custom table-valued functions](../sql/custom_tabular). With this mechanism, we can expose data from other databases to our SQLite engine. This chapter explains how.

> To see an example of cross-database querying, download the [sample workbook](../demofiles/remote_db_query.xlsx).

## 1. Define an embedded query to the external database

Suppose we have a SQL Server database and we want to expose a table called `Departments` to SQLite. The first step is to define a query that returns the data we're interested in. 

```sql
select * from HumanResources.Department d
``` 

In order to expose this data to SQLite, we need to [embed this query into the workbook](../automation/embedding_code.md). I'll give the query a name: *Get departments*.

![embedding query](../images/emb_dpt.png)

## 2. Call the embedded query from C# #

Now that the query is embedded, we can run it from C# like so:

```csharp
Query("Get departments").Run().Items
```

![calling embedded query from C#](../images/cs_emb_query.png)

The above script will run the query and return the results as a `dynamic[]` object. We can expose this data to SQLite as a table-valued function, but we first need to map the rows to a specific type, since QueryStorm needs to know what columns the function will return (i.e. we can't return dynamically typed objects). 

We could write the new type and the mapping code ourselves, but there's a handy context menu command that does it for us:

![generate api command](../images/generate_api_command.png)

This command will run the query in the background to determine which columns are returned. Based on that information, it generates a strongly typed class that represents rows, and a function that calls the query and exposes the results to SQLite. It embeds this code in a script in the `APIs` folder, and gives it a name based on the name of the original query. 

![generate api result](../images/generate_api_result.png)

## 3. Call the query from SQLite

We can now disconnect from our current connection and reconnect using SQLite. Once we do, we'll see the new table-valued function in autocomplete, and we'll be able to use it:

![generate api result](../images/remote_db_tvf.png) 

The above query is executed by SQLite, but the table valued function internally loads data from a SQL Server database.

We can repeat the same procedure with multiple embedded queries. This will make data from multiple databases available to SQLite. If we like, we can join workbook tables with data from web services and multiple external databases and then dump the results into Excel or a [SQLite file database](../sql/sqlite_files).

## Limitations
Currently, it is not possible to pass parameters to embedded queries. Allowing this is planned in a future release. Once ready, it will enable selectively loading data from external databases which will minimize the amount of data that needs to be loaded when e.g. running joins with data from external databases.  