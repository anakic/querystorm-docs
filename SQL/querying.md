# SQL in Excel
QueryStorm comes with a [SQLite](https://www.sqlite.org) database engine that can work with Excel tables as if they were database tables.  

!!! Data
	 In this example, I use a [public dataset](https://www.querystorm.com/downloads/demos\movies.xlsx) with data on 5000 movies scraped from IMDB.

## Connecting
Clicking the connect button pops up QueryStorm and opens a SQLite connection that *sees* Excel tables as database tables.

![Connect to workbook](https://i.imgur.com/pT3C7tq.png)

## Querying

Once connected, we can start querying. The most important thing to know is that QueryStorm works with [Excel tables](http://www.excel-easy.com/data-analysis/tables.html "Excel tables intro"), **not sheets**.

![Querying](https://i.imgur.com/YsnFY0g.png)


### Selecting rows
Let's go through a few simple examples. For a start, let's analyze the budgets and incomes of movies that are about the future:

```sql
select avg(budget), avg(gross), avg(gross)/avg(budget) as return
from movies
where plot_keywords like '%future%'
``` 
![Select example #1](https://i.imgur.com/Fxr3Bgy.png)

It looks like movies about the future are good business; they return 115% of the budget spent.

Now let's look at how incomes and scores of movies changed over decades:

``` SQL
select
	cast (title_year / 10 as int) * 10 as decade, 
	count(*) [nr. of movies], 
	round(avg(imdb_score), 2) [avg score], 
	Format('{0:#,##0.00}', avg(budget)) [avg budget], 
	Format('{0:#,##0.00}', avg(gross)) [avg gross],
	round(avg(gross)/avg(budget), 2) as [return]
from
	movies
where
	title_year is not null
group by
	decade
order by
	decade
```   
![Movies breakdown by decade](../images/movies_by_decade.png)

We can now proceed to draw some overgeneralized conclusions from our sample data: 

- IMDB scores appear to be higher for older movies
- in 2000's spending in the movie industry peaked, but wasn't profitable
- return on investment in 1930's was a whopping 42x, but has been declining ever since

### Updating rows
We're not limited to just querying, we can also modify the data in Excel tables. Let's put a star next to the highest grossing movie of each director:

![Update example](https://i.imgur.com/XRh4h5K.png)

In the above query, I'm looking for all movies such that no movie from the same director grossed more. There were 2695 rows updated, one per director.  
 
### Deleting rows
For some movies, we don't have information about the gross income they produced. We're analyzing income, and these rows are in the way, so let's just delete them:

``` sql
delete from movies where gross is null
```
The *messages* pane indicates that 772 rows have been deleted:
```
Command executed in 220ms, 0 rows returned, 772 rows affected
```

### Inserting rows
Lastly, we can also insert data into tables. For example, let's insert a row and supply the director name and movie title. We'll leave the other columns empty:

``` sql
insert into movies(director_name, movie_title)
values ('Jeff Jackson', 'Sound plan 8')
```  
![Row inserted](../images/row_inserted_sql.png)

We can also insert the results of a select statement to insert a bunch of rows at once.

```sql
insert into movies(director_name, movie_title)
select director, movie from newMovies
```

## Special columns
When using the SQLite engine, all workbook tables get two extra hidden columns: `__address` and `__row`. These columns are *hidden*, meaning that SQLite will not include them in the results if you just specify `*` in the select list, but you can include them explicitly in the select list when you need them.  

![Data in a sheet](https://www.querystorm.com/Images/docs/special_columns.png)

### The `__address` column 
The `__address` column is for locating the original row in Excel. Double-clicking an address in the results (or the row header, if an address is present in the row) will select the row in Excel. We can also use it to [perform formatting operations](./formatting) on the row.

### The `__row` column
The `__row` column specifies the original index of the row in Excel. We can use this to preserve the original ordering of rows. It's also convenient for some other operations, e.g. self-joins without duplicates.