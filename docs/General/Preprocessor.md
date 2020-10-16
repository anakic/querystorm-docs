# The SQL Preprocessor

QueryStorm introduces a simple preprocessor that's available in all SQL scripts in QueryStorm. The job of the preprocessor is to allow:

- Using values from workbook cells in SQL code
- Defining Excel functions
- Defining Excel commands

## Inserting values into SQL

The preprocessor can be used to insert values from Excel cells into SQL code, for example:

```sql
select * from people where id = {nameOfCell}
```

The value of the named cell will be inserted instead of the `{nameOfCell}` placeholder before the query is passed to the SQL engine. This is a purely textual operation. If the inserted value is a string or a date, it is automatically quoted.

To prevent quoting, add an equals sign before the name of the cell: `{=nameOfCell}`. This allows inserting raw SQL code into the query (SQL injection), so caution is advised when doing this.

### User defined variables

Users can also define variables in code. This is especially useful when working with database engines that don't natively support defining variables (e.g. SQLite, Access).

A variable can be defined and referenced in the following way:

```sql
-- define the variable
{var binSize = 20}

SELECT
    CAST (totalPay / {binSize} AS int) * {binSize} AS bin, count(*) AS count
FROM
    salaries
GROUP BY
    bin
```

### Formatting values before insertion

Values can be formatted before inserting them into the query. The syntax for formatting is as follows: {*variableName*|*formatSpecifier*}.

The format specifier is a standard format specifier for the .NET [string.Format](https://docs.microsoft.com/en-us/dotnet/api/system.string.format) method.

For example:

```sql
{datetime myDate = "2020-10-10"}

SELECT
    {myDate|"The date {0:d} is a {0:dddd}"}

-- the output will be: The date 10/10/2020 is a Saturday
```

### Preprocessor expressions vs parameters

SQL scripts can reference named cells as parameters (if they support named parameters), so inserting values into the query text with the preprocessor usually isn't required.

However, there are several reasons why this might be useful:

- It allows defining and referencing variables with engines that do not natively support user-defined variables (e.g. SQLite)
- It allows formatting values before inserting into SQL code
- It allows using cell values as parameters when working with databases that do not support named parameters (e.g. ODBC, Access)
- It allows injecting raw SQL into queries

## Defining Excel functions

The second important task of the preprocessor is declaring functions that can be used from Excel. The body of the function is written in SQL and the declaration of the function via the preprocessor.

For example:

```sql
{function myFunction(int rowsToReturn = 20)}

select top {rowsToReturn}
    *
from
    HumanResources.Department d
```

The above query can be used to define an Excel function that accepts a single parameter and returns a table of results. For more information on creating functions with SQL, click [here](../../Functions/Functions_via_SQL).

## Defining Excel commands

The preprocessor also allows defining commands. SQL code is used for fetching data or modifying data in the database, while the preprocessor syntax is used to declare the command and specify the events that will trigger the execution of the command.

For example:

```sql
{handles Sheet1!CommandButton1}

insert into PermanentDbTable from ImportedWorkbookTable
```

If the query returns data from the database (rather than updating the database), the query results can be written into the workbook by adding an output directive above it, for example:

```sql
{handles Sheet1!CommandButton1}

{@OverdueBooks}
select
    BookName, PersonName
from
     BookRentals
where
    ReturnDate < date()
    and Returned=0
```

Running this query will immediately update the *OverdueBooks* table in Excel with the results.

For more information on setting up automation with SQL, click [here](../../Automation/Automation_via_SQL).
