# Parameterizing queries 
Some times we might want to use values from Excel cells as parameters in our queries. This can be especially useful when automating queries. 

This is supported via QueryStorm's [SQL preprocessor](https://www.querystorm.com/docs/Preprocessor "Preprocessor"):

``` SQL
select * from movies where gross > {Sheet1!A1}
```

Cells can be referenced in the following ways:

1. By address, e.g. `{A1}`
2. By qualified address e.g. `{Sheet1!A1}`
3. By named range e.g. `{myNamedRange}`

Option 3 (named ranges) is recommended, as it enables autocomplete, the cell can be moved without breaking queries and there is no chance of reading the value from the wrong sheet.

!!! Note 
	Quotes are added automatically if the preprocessor expression resolves to a string or a date, so you should not add them yourself. 