# Formatting via SQL

Since the SQLite engine is running *in-process* with Excel, it can interact with Excel objects. A typical use case for this is modifying formatting. 

## Formatting rows

As [mentioned before](./querying/#special-columns), each table has a hidden `__address` column which we can use to selectively apply formatting to rows.

Let's set the background color for movies that grossed more than $40M:
``` SQL
--first clear any previous formatting 
select ClearBackgroundColor('movies');

--apply new formatting
select
	*, SetBackgroundColor(__address, 'Orange')
from
	 movies
where
	gross > 400000000
order by
	__row asc
``` 
![Formatting rows example](https://i.imgur.com/VVGYJ68.png)

In the example, we're using the `ClearBackgroundColor` function to clean any previously applied formatting, after which we apply formatting to the desired rows via the `SetBackgroundColor` function.

It might seem strange to use a `SELECT` statement to modify row formatting, but it's fairly easy to get used to, and there doesn't seem to be a more elegant way of doing it via SQL.

## Formatting cells

We can use the `SetBackgroundColor` function to format cells as well. Here's an example of coloring cells based on their values:

``` SQL
select
	*,
	(case
		when value > 100000 
			then SetBackgroundColor(Address, 'blue')
		else 
			SetBackgroundColor(Address, 'red')
	end) color
from
	xlcells()
where 
	type = 'Double'
```

![Format cells](../images/format_cells.png)