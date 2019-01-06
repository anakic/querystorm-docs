# Parameterizing queries 
It can be handy to use cells as parameters in queries, especially within automated queries. This is enabled by [preprocessor expressions](../misc/preprocessor "Preprocessor"). 

``` SQL
select * from movies where gross > {Sheet1!A1}
```

Cells can be referenced in the following ways:

1. By address, e.g. `{A1}`
2. By qualified address e.g. `{Sheet1!A1}`
3. By named range e.g. `{myNamedRange}`

!!! Tip
	Option 3 (named ranges) is recommended in embedded queries because the cell can be moved without breaking queries and there is no chance of reading the value from the wrong sheet.

The expression in curly brackets gets evaluated and replaced by its result and the resulting query is sent to the database engine. Quotes are added automatically if the expression resolves to a string or a date, so you should not explicitly add quotes around expressions.