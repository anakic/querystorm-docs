# SQL in Excel
QueryStorm comes with a built-in [SQLite](https://www.sqlite.org) database engine that enables using SQL on Excel tables. 

> The examples in this article use a [public dataset ](https://www.querystorm.com/downloads/demos\movies.xlsx) containing data on 5000 movies scraped from IMDB.

## Connecting
Clicking the connect button creates an in-memory SQLite database which *sees* Excel tables as if they were database tables, and opens up the QueryStorm IDE.

![Connect to workbook](https://i.imgur.com/pT3C7tq.png)

## Querying

Once connected, we can start querying. The most important thing to know for a start is that **QueryStorm works with Excel tables** [[not to be confused with sheets](http://www.excel-easy.com/data-analysis/tables.html "Excel tables intro")].

![Querying](https://i.imgur.com/YsnFY0g.png)


### Selecting rows
Let's run a few queries. For a start, let's analyze the budgets and incomes of movies that are about the future:

```sql
select avg(budget), avg(gross), avg(gross)/avg(budget) as return
from movies
where plot_keywords like '%future%'
``` 
![Select example #1](https://i.imgur.com/Fxr3Bgy.png)

It turns out making a movie about the future is good business; they seem to return 115% of the budget spent.

That was pretty simple. For a slightly more complex example, let's take a look at how the incomes and scores of movies changed through the decades:

``` SQL
select
	cast (title_year / 10 as int) * 10 as bin,
	count(*) count,
	avg(imdb_score) score,
	avg(budget),
	avg(gross) gross
from
	movies
group by
	bin
order by 
	bin
```   
![Select example #2](https://i.imgur.com/FHSKBLr.png)

Since we're here, let's draw some (overly-general) conclusions (from insufficient data): 

- IMDB scores appear to be higher for older movies
- in 2000's spending in the movie industry peaked, but wasn't profitable
- return on investment in 1930's was a whopping 40x, but has been declining ever since

### Updating rows
The SQLite engine is not limited to querying. We can also modify data in Excel tables using SQL `update` statements. Let's update all rows and put a star next to the highest grossing movie of each director:

![Update example](https://i.imgur.com/XRh4h5K.png)

In our query, we're looking for all movies such that no movie from the same director grossed more. There were 2695 rows updated (one for each director), meaning that on average there are ~2 movies per director in the list (4926 movies / 2695 directors).  
 
### Deleting rows
For some movies, we don't have information about the gross income they produced. We don't care about those, so let's just delete them:

``` sql
delete from movies where gross is null
```
The *messages* pane informs us that 772 rows have been deleted:
```
Command executed in 220ms, 0 rows returned, 772 rows affected
```

### Inserting rows
Lastly, we can also insert data into tables. Let's insert a row and supply just the director and title:

``` sql
insert into movies(director_name, movie_title)
values ('Jeff Jackson', 'Sound plan 8')
```  
> The row is inserted at the end of the table.

If we like, we can also insert the results of a select statement to bulk insert many rows at once.

```sql
insert into movies(director_name, movie_title)
select director, movie from newMovies
```

## Special columns
When using the SQLite engine, all workbook tables get two extra hidden columns: `__address` and `__row`. These columns are *hidden* by default, meaning that SQLite will not include them in the results if you just specify `*` in the select list, but you can include them explicitly in the select list when you need them.  

![Data in a sheet](https://www.querystorm.com/Images/docs/special_columns.png)

### The `__address` column 
The `__address` column is useful for locating the original row in Excel. Double clicking it (or the row header, if an address is present in the row) in the results grid will select the row in Excel. Additionally, we can also use the address in order to [perform formatting operations](./formatting) on the row.

### The `__row` column
The `__row` column specifies the original index of the row in Excel. We can use this to preserve the original ordering of rows (and some other things e.g. self-joins without duplicates).
