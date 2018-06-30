# Multi-database queries

The SQLite engine in QueryStorm access data in external databases using [custom table-valued functions](../sql/custom_tabular). Basically the idea is to call custom table-valued functions (defined in C#) that call embedded queries that load data from external databases. Here's how to do it...

> You can download the sample workbook from [here](https://www.querystorm.com/downloads/demos/remote_db_query.xlsx). It does not have any data in the sheets, but it has the embedded queries that illustrate the concepts.

## 1. Define an embedded query to the external database

Suppose we have a SQL Server database and we want to expose a table called `Departments` to SQLite. The first step is to define a query that returns the data we're interested in. 

```sql
select * from HumanResources.Department d
``` 

In order to be able to expose this data to SQLite, we need to [embed this query into the workbook](../automation/embedded_queries.md). Queries are embedded along with the connection configuration. Let's name the embedded query "Get departments".

## 2. Call the embedded query from C# #

Now that the query is embedded, we can run it from C# like so:

```csharp
Query("Get departments").Run().Items
```

![calling embedded query from C#](https://www.querystorm.com/downloads/images/cs_emb_query.png)

The above script will run the query and return the results as a `dynamic[]` object. We can expose this data to SQLite as a table-valued function. However, we first need to define a type for the rows, as we can't expose rows of type dynamic. This means we need to define a custom type that will represent the row of data and map the dynamic objects to the new type. We could write this code ourselves, but there's a handy context menu command that does it for us:

![generate api command](https://www.querystorm.com/downloads/images/generate_api_command.png)

This command will run the query in the background to determine which columns are returned. Based on that information, it generates a strongly typed class that represents rows, and a function that calls the query and exposes the results to SQLite. It embeds this code in a script in the `APIs` folder, and gives it a name based on the name of the original query. 

![generate api result](https://www.querystorm.com/downloads/images/generate_api_result.png)

> If you need to modify the script by hand, it's recommended to first disconnect from the remote database and connect using C#, as this will make auto-complete and error checking available while you're modifying the code.

## 3. Call the query from SQLite

We can now disconnect (from the remote database or from C#) and reconnect using SQLite. Once we do, we'll see the new table-valued function in autocomplete, and we'll be able to use it:

![generate api result](https://www.querystorm.com/downloads/images/remote_db_tvf.png) 

We can repeat the same procedure with multiple embedded queries. This will make data from multiple databases available to SQLite so, for example, we can have a SQLite query that joins excel tables with data from several external databases plus data from a web service.

## Limitations
Currently it is not possible to pass parameters to embedded queries. This is not yet implemented but is planned. Once ready, it will enable selectively loading data from external databases which will minimize the amount of data that needs to be loaded when e.g. running joins with data from external databases.  