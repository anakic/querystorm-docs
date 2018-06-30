# Formatting via SQL

Since the SQLite engine is running *in-process* with Excel, it can interact with it in many ways e.g. it can modify formatting in Excel. 

## Formatting rows

As [mentioned before](./querying/#special-columns), each row has a hidden `__address` field, and we can use to specify where the formatting needs to be applied. 

Let's set the background color for movies that grossed more than $40M:
``` SQL
--clear any previous formatting 
select ClearBackgroundColor('movies');

--apply formatting
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

It might seem strange to use a `SELECT` statement to modify row formatting, but it's fairly easy to get used to, and there doesn't seem to be a better way of doing it via SQL.

## Formatting cells

As with rows, we can use the `SetBackgroundColor` function to format cells.

Let's color cells based on the values they contain:

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
	cells ()
where 
	type = 'Double'
```

Here we're formatting all cells that contain a number in one of two colors, depending on their value. Since the selection might contains strings and dates, we can use the where clause to filter them out to not modify their formatting.

![Formatting cells example](https://i.imgur.com/xWT1Pgr.png)