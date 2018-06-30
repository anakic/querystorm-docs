# SQLite in Excel
QueryStorm comes with a built-in SQLite database engine that enables using SQL in Excel. This article describes what it can do and gives some examples of how to use it.

> The examples in this article use a public dataset containing data on 5000 movies scraped from IMDB. Download it from [here](https://www.querystorm.com/downloads/demos\movies.xlsx) to follow along.

## Connecting
To connect using the SQLite engine, click the button in the ribbon. This creates a new in-memory SQLite database which uses an adapter to attach to all Excel tables in the current workbook. 

![Connect to workbook](https://i.imgur.com/pT3C7tq.png)

> You can use the button's dropdown to use a different scope, i.e. current sheet or all open workbooks.  

## Querying

Once connected, we can start writing queries. An important thing to know is that **QueryStorm works with Excel tables** ([not to be confused with sheets](http://www.excel-easy.com/data-analysis/tables.html "Excel tables intro")).

![Querying](https://i.imgur.com/YsnFY0g.png)

Let's see some examples...

### Selecting rows
Let's compare the average budget and income for all movies with the word 'future' in the plot keywords:

```SQL
SELECT avg(budget), avg(gross), avg(gross)/avg(budget) as return
FROM movies
where plot_keywords like '%future%'
``` 
![Select example #1](https://i.imgur.com/Fxr3Bgy.png)

As it turns out, making a movie about the future is good business; they seem to return 115% of the budget spent.

For a slightly more complex example, let's compare the incomes and scores of movies through the decades:

``` SQL
SELECT
	cast (title_year / 10 as int) * 10 as bin,
	count(*) count,
	avg(imdb_score) score,
	avg(budget),
	avg(gross) gross
FROM
	movies
group by
	bin
order by 
	bin
```   
![Select example #2](https://i.imgur.com/FHSKBLr.png)

We can draw some interesting conclusions based on this data: 

- IMDB scores appear to be higher for older movies
- in 2000's spending in the movie industry peaked, but wasn't profitable
- return on investment in 1930's was a whopping 40x, but has been declining ever since

> Obviously, this is a small sample (5000 movies) with relatively little data on old movies, so in practice we'd have to be a bit more careful about declaring these results conclusive. 

### Updating rows
The SQLite engine is not limited to querying. We can also modify data in Excel tables just as if they were regular database tables.

Let's update all rows and put a star next to the highest grossing movie of each director:

![Update example](https://i.imgur.com/XRh4h5K.png)

In the query we're looking for all movies, such that there isn't another movie from the same director that grossed more. There were 2695 rows updated, one for each director, meaning that an average director has ~2 movies in the list (4926 / 2695).  
 
### Deleting rows
For some movies we don't have information about the gross income they produced. Rather than guarding against this in our queries, let's delete these rows:

``` sql
delete from movies where gross is null
```
The *messages* pane informs us that 772 rows have been deleted:
```
Command executed in 220ms, 0 rows returned, 772 rows affected
```

### Inserting rows
Lastly, we can also insert data into tables, although this is usually less useful than the other operations. Let's insert a row and supply just the director and title:

``` sql
insert into movies(director_name, movie_title)
values ('Jeff Jackson', 'Sound plan 8')
```  
> The row is inserted at the end of the table.

### Hidden columns
When using the SQLite engine, all workbook tables get two extra hidden columns: `__address` and `__row`. These columns are *hidden* by default, meaning that SQLite will not include them in the results if you just specify `*` in the select list, but you can include them explicitly in the select list when you need them.  

![Data in a sheet](https://www.querystorm.com/Images/docs/special_columns.png)

#### The `__address` column 
is useful for locating the original row in Excel. Double clicking a cell in the results grid that contains an Excel address will select the range in Excel! Additionally, we can also use the address in order to perform formatting operations on the row, as we'll see in the next section of this article.

#### The `__row` column
specifies the original index of the row in Excel, which can come in handy in all sorts of situations (preserving the original ordering, creating self-joins without duplicates etc).

## Formatting via SQL
Since the SQLite engine is running *in-process* with Excel, it can interact with it in other ways, such as modifying formatting for rows using the `__address` column and functions for interacting with Excel.

One such function is `SetBackgroundColor`.

Let's use it to set the background color movies that grossed more than $40M:
``` SQL
--clear any previous formatting 
select ClearBackgroundColor('movies');

--apply formatting
SELECT
	*, SetBackgroundColor(__address, 'Orange')
FROM
	 movies
where
	gross > 400000000
order by
	__row asc
``` 
![Formatting rows example](https://i.imgur.com/VVGYJ68.png)

> In the example, we're using the `ClearBackgroundColor` function to clean any previously applied formatting.

## Working with individual cells

Aside from working with tables, QueryStorm's SQLite engine can also work with individual cells. This is useful when working with unstructured data.

To work with cells, we use the **`xlCells`** table-valued function.

Return a list of cells in the current selection:
```sql
select * from xlcells()
``` 
Return a list of cells in the specified range
``` sql
select * from xlcells('Sheet1!B5:D9')
```
Here's what the returned data looks like:
![Cells query](https://i.imgur.com/iHcra7e.png)
 
For each cell, the function returns a row in the results that has the follwing columns:
-  Address
-  Row
-  Column
-  Column letter
-  Value
-  Formula

### Updating cells
We can't use a SQL `update` statement on `xlCells` to update cell values, but we CAN update them using the `SetCellValue` function.

For example, let's reverse the text in the selected cells:
``` SQL
SELECT
	*, 
	SetCellValue(Address, reverse(Value)) --get the value of the cell, reverse it and write it back into the cell
FROM
	xlcells ()
```
![Update cell values example](https://i.imgur.com/XbDgKfu.png)

> The selection can be supplied explicitly via addres, or via selection. The selection can consist of one or more areas (Ctrl+Select in Excel) 


### Formatting cells
As with rows, we can use the `SetBackgroundColor` function to format cells.

Let's color cells based on the values they contain:

``` SQL
SELECT
	*,
	(case
		when value > 100000 
			then SetBackgroundColor(Address, 'blue')
		else 
			SetBackgroundColor(Address, 'red')
	end) color
FROM
	xlcells ()
```
![Formatting cells example](https://i.imgur.com/xWT1Pgr.png)

## SQLite file dabases

### Why use a file database?
Up until now, we've been working with Excel tables in-memory. Another important thing the SQLite engine can do is connect to (or create) .sqlite database files. This enables you to consume data from existing .sqlite files, but it also enables you to use a .sqlite file as a data store, instead of keeping the data in your workbook.

For example, if you have a large Excel workbook, you could dump the data from the workbook into a .sqlite file, and use Excel for user input, displaying results and visualizations. Since the data is stored in a .sqlite file and not the workbook itself, you don't have to wait a long time every time you open the Excel file, and Excel generally behaves better with less data.

In effect, it enables you to use a sqlite file for storage and Excel as the user interface.

### Connect to a sqlite file 

To connect to a sqlite file, we must define a connection. We can choose which sqlite file to connect to, and also which Excel tables to include in the connection. We can define a new connection via the ***Connect with*** button in the ribbon, which opens up the following dialog:

![Create new connection](https://i.imgur.com/bsxnWn1.png)

> Description of the steps:
> 1. Create new connection
> 2. Choose SQLite as the engine
> 3. Enter the connection string manually(3a) or via the connection string editor(3b)
> 4. Set the scope, or select tables manually
> 5. Press Connect

You can select an existing .sqlite file, or you can create a new one by specifying a file name that does not yet exist.

Once connected, you can query permanent tables, as well as Excel tables.

### Moving data into a SQLite file

Let's move the `movies` table from Excel into the sqlite file.

```sql
create table movies as -- create permanent table 'movies'
select * from movies -- using data from in-memory table 'movies'
```
This might seem like a name collision, but the in-memory table's full name is `temp.movies` and the name of the new permanent table is `main.movies`. If both tables exist and have the same name, we can refer to the permanent table like so:
```sql
select * from main.movies
``` 
More importantly, we can now delete the table from Excel or connect to the .sqlite file from a brand new Excel file. Rather than keeping the data in the workbook, we use the workbook to present results and graphs to the user and use the .sqlite file for storage.

### Moving data into Excel
All SQL engines in QueryStorm can use the preprocessor to redirect the results of a query into an Excel table or range.

For example, we can write the contents of the `actors` database table into a new or existing Excel table called `act` like so:

```sql
@Table TableName="act" IfExists="Update" DefaultAddress="B2"
select * from actors
```

You can read more about the preprocessor [here](https://www.querystorm.com/docs/Preprocessor "Preprocessor")

## Multiple workbooks

You can connect to tables from multiple workbooks in two ways:

- Via the dropdown button in the ribbon
- Via custom connection

This allows you to run queries that span tables from multiple workbooks. The workbooks have to be opened in order for this to work, so this is primarily useful for ad hoc purposes rather than automation.

## Parameterized queries 
Often times it can be useful to use the values in Excel cells as query parameters. This is possible using QueryStorm's SQL preprocessor. 

Here's an example:
``` SQL
select * from movies where gross > {Sheet1!A1}
``` 

You can read more about the preprocessor [here](https://www.querystorm.com/docs/Preprocessor "Preprocessor"). 

## Functions
SQLite comes with many useful funcitons built in, but QueryStorm also registers many additional functions of its own into the SQLite engine. This includes functions for working with regular expressions, interacting with Excel and other specialized tasks (running VBA functions, calculating GPS distance etc...).

You can see a list of scalar and tabular functions in the object explorer:

![Functions](https://i.imgur.com/pX2OtZF.png)

You can view the documentation for a given function in autocomplete, in tooltips (object explorer and editor) and in the [function reference](https://www.querystorm.com/docs/SQLite_functions "function reference"). 

> There's a subtle naming convention here: functions that start with an uppercase letter are QueryStorm's expansions to the database engine.