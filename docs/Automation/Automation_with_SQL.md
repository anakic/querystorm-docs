# SQL commands

Automation in QueryStorm is usually implemented by writing C# or VB.NET code inside component classes, as described in [the previous chapter](../Automation_with_dotnet). However, this code does not necessarily need to be written by hand. SQL scripts can generate the component code automatically, which is especially useful for SQL users who are not familiar with C# and VB.NET. These scripts are called **SQL commands**.

For this reason (among others) QueryStorm introduces a [preprocessor syntax](../../General/Preprocessor) for SQL. With SQL commands, SQL code is used to fetch/update data while preprocessor code is used to specify when the script should be executed, and where the results should be written to (e.g. an Excel table).

## Example

Let's suppose we want to populate an Excel table with sales orders for a given date. When the value of the cell that contains the date changes, we want the script to execute again and to output the new results into the table.

Here's how we might do that:

```sql
-- specify the triggers that call the command
{handles orderDate}

-- output results into the 'orders' table
{@orders}
select
    *
from
    Sales.SalesOrderHeader soh
where
    OrderDate = @orderDate -- read the orderDate cell's value
```

## Starting the application

In order to start the application, the script needs to be saved (++ctrl+s++), and the project built (++ctrl+shift+b++). The component code is generated when the script is saved. After the build completes, the runtime loads and activates the workbook application.

## Accessing workbook tables and variables

All cells with assigned names are visible inside scripts as parameters. The code in the previous example uses a cell called `orderDate` as a parameter.

To allow the script to access Excel tables as well, the tables must be included when configuring the script via the "Connect" dialog.

## Outputting query results

The preprocessor allows sending query results into Excel (via the data context). In the example above, the `{@orders}` output directive is used to send the results of the `select` query into an Excel table called `orders`.

The syntax of the output directive is very simple: {@*table_name*}

The output directive should be placed above the select query whose results it should output. If there are multiple select queries in the script, each of them can have its own output directive, so multiple tables can be updated from the same script.

## Specifying triggers

The preprocessor syntax for declaring an automation command is: {**handles** *eventsList*}. The event list is a comma-separated list of events that should trigger the execution of the command.

### Range change events

For each named cell, an event with the same name is defined and is fired each time the cell value changes. In the above example, the `orderDate` event is specified as the only trigger, meaning that the command will execute every time the `orderDate` cell's value changes.

### ActiveX button click events

To handle the click of an ActiveX button, we should use the following syntax: {**handles** *sheetName*!*buttonName*}, for example `{handles Sheet!CommandButton1}`.

### VBA events

Arbitrarily named events can also be sent from VBA and used to trigger the execution of commands. To send an event from VBA, use the QueryStorm.Runtime.API class:

```vb
CreateObject("QueryStorm.Runtime.API").SendEvent("myEvent")
```

The event can be handled using the preprocessor like so: `{handles myEvent}`

Scripts can handle multiple events. Event names should be separated by a comma:

```sql
{handles myEvent, orderDate, Sheet1!CommandButton1}

{@orders}
select
    ...
```

## Commands vs Functions

The preprocessor supports defining commands as well as [functions](../../Functions/Functions_via_SQL). The command and function declarations are mutually exclusive. A script can either declare a command or a function, but not both.

Example function declaration:

```sql
{function myFunc(int param1=1, string param2="abc")}
// ...sql code
```

Example command declaration:

```sql
{handles myNamedCell, Sheet1!MyButton}
// ...sql code
```

In general, functions fetch or calculate values while commands change things (e.g. save data to a database, write data into Excel tables).

In the example at the beginning of this page, we were fetching data based on a parameter. Conceptually, this looks more like it should be a function than a command. However, a function cannot return the body of a table, it can only return data to a cell or a range. If we want to write the resulting data to an Excel table, we need to use a command with an output directive (e.g. `{@someTable}`).

On the other hand, if the script was **updating database data** rather than simply selecting data from it, a command would certainly be the appropriate choice.

An important distinction between commands and functions is that commands reset Excel's Undo/Redo stack when they modify data in the workbook, while functions do not. This is a technical limitation in Excel.
