# SQL commands

Automation in QueryStorm is usually implemented by writing C# or VB.NET code inside component classes, as described in the previous chapters. However, this code does not necessarily need to be written by hand. SQL scripts can generate the component code automatically, which is especially useful for SQL users who are not familiar with C# and VB.NET. These scripts are called **SQL commands**.

For this reason (among others) QueryStorm introduces an additional [preprocessor](todo) syntax for SQL scripts. With SQL commands, SQL code is used to fetch/update data and preprocessor code is used to specify when the script should be executed, and where the results should be written to (e.g. an Excel table).

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

In order to start the application, the script needs to be saved (`Ctrl+S`), and the project built (`Ctrl+Shift+B`). The component code is generated when the script is saved, and after the build, the runtime loads and activates the workbook application.

## Accessing workbook data

The code in the example assumes that a cell called `orderDate` is defined in the workbook. The variable `@orderDate` that refers to this cell is made available to the script via the data context.

Aside from giving cells names, no other special handling is required in order to be able to reference their values inside scripts.

To allow the workbook to access Excel tables as well, the tables must be included when setting up the script via the "Connect" dialog.

## Outputting query results

The preprocessor allows sending query results into Excel (via the data context). In the example above, the `{@orders}` output directive is used to send the results of the `select` query into an Excel table called `orders`.

The syntax of the output directive is very simple: {@*table_name*}

The output directive should be placed above the select query whose results it should output. If there are multiple select queries in the script, each of them can have its own output directive, so multiple tables can be updated from the same script.

## Specifying triggers

The preprocessor syntax for declaring an automation command is: {**handles** *eventsList*}. The event list is a comma separated list of events that should trigger execution of the command.

### Range change events

For each named cell, an event with the same name is defined, and is fired each time the cell value changes. In the initial example, the `orderDate` event is specified as the only trigger, meaning that the command will execute every time the `orderDate` cell's value changes.

### ActiveX button click events

To handle the click of an ActiveX button, we should use the following syntax: {**handles** *sheetName*!*buttonName*}, for example `{handles Sheet!CommandButton1}`.

### VBA events

Arbitrarily named events can also be sent from VBA and used to trigger execution of commands. To send an event from VBA, use the QueryStorm.Runtime.API class:

```vb
CreateObject("QueryStorm.Runtime.API").SendEvent("myEvent")
```

The event can be handled using the preprocessor like so: `{handles myEvent}`

The script can handle multiple events. Event names should be separated by a comma:

```sql
{handles myEvent, orderDate, Sheet1!CommandButton1}

{@orders}
select
    ...
```

## Commands vs Functions

The preprocessor supports defining commands as well as [functions](todo). The command and function declarations are mutually exclusive. A script can either declare a command or a function, but not both.

The functionality in the example looks like it should be implemented as a function that would accept a date parameter and return the query results. If the user's version of Excel supported dynamic arrays, the results could even automatically spill. Additionally, functions don't mess with Excel's undo functionality, which is another advantage for functions.

The drawback of functions, however, is that the data returned from a function isn't an Excel table. The function results cannot be used as the contents of a table, so if you need to do any further processing on the data, the table is a much better place to store the results. **Writing to a table** can only be done from a command.

For this example, since we're returning data from a database, we could have implemented this functionality either as a function or a command.

However, if the script was **updating database data** rather than simply selecting data from it, implementing it as a command would be the appropriate choice.
