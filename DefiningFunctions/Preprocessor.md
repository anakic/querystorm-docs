# The SQL Preprocessor

To allow SQL scripts to do more than just querying and modifying data, QueryStorm introduces a simple preprocessor that's available in all SQL scripts in QueryStorm. 

Among other things, the preprocessor allows you to: 
- Use values from Excel as parameters
- Define Excel functions
- Output results into Excel
- Trigger commands by Excel events

## Parameterizing queries



### Referencing cells 

For example, let's assume cell *A1* in *Sheet1* contains the value 3:

```sql
select * from people where id = {Sheet1!A1}
```
The following query will be passed on to the underlying database engine:
```sql
select * from people where id = 3     
```

Here's how QueryStorm expressions are processed:
1. Evaluate expression
1. If result value is a string, surround it with single quotes
1. Replace the preprocessor expression in the query with the results
1. Once all expressions are processed, pass the updated query on to the underlying engine 

### Preprocessor variables

Additionally, it's also possible to define variables inside queries. This helps reduce *magic constants* and repetition, as well as making it easier to run the query with different inputs.

Here's an example:

```sql
:{myVariable = 123}

select * from people where id = {myVariable}
```
The following is sent to the underlying database engine:
```sql
select * from people where id = 123
```


## Redirecting results
The second important task of the preprocessor is redirecting query results into Excel tables and ranges. This can come in very handy for automation. For this purpose QueryStorm's preprocessor provides "*@-directives*".

Here's an example:


```sql
@Table TableName="OverdueBooks" IfExists="Update"
select BookName, PersonName from BookRentals where ReturnDate < date() and Returned=0 
```

Running this query will immediately update the *OverdueBooks* table in Excel with the results. This saves time if we're working on a query and always writing the results into the same Excel table, but more importantly, it enables automatic update of an Excel table if we embed the query and set up appropriate triggers for executing it (automation). 

### Multiple queries and redirects
It is possible to have multiple select statements within a query, each one outputting results into a different table or range.

Example:

```sql
@Table TableName="LargestCities"
select * from cities order by population descending limit 10

@Table TableName="SmallestCities"
select * from cities order by population limit 10 
```

Running this query will update two excel tables (*LargestCities* and *SmallestCities*) in one go. For automation scenarios, this can save time and bandwidth as the same connection (session) can be used to update multiple tables.

 
   