# The IDE

This section explains the main segments of the IDE and gives a few useful tips&tricks to get the most out of the IDE. 

![IDE screenshot](http://www.querystorm.com/Images/Docs/ide.png?v=1)

## 1. Code Editor
QueryStorm's code editor is equipped with parsers which it uses to analyze your code as you're typing it. The parsers enable it to understand your code so it can offer meaningful information, hints and code completion suggestions.

Parsers currently exist for SQLite and TSQL dialects, but there is currently no parser support for the other dialects. As a fallback, the other engines use the SQLite parser with error highlighting disabled (to avoid false error reports due to dialect differences).

> Tip: adjust the font size using `Ctrl`+`mouse scroll`.

### Error highlighting 
On each keystroke, the parser analyzes the query and uses the database schema information to detect possible errors in your code. The errors are underlined in the editor but will not prevent the execution of the query (just in case the error is a *false positive*). 

![Error highlighting](http://www.querystorm.com/Images/Docs/errorhighlighting.png)

The parser can detect the following types of errors:

- syntax errors (badly formed query)
- undefined symbols (referring to objects that are not defined or are out of scope)
- ambiguous symbols (e.g. a column that can come from two or more tables present in the `FROM` clause)
- certain semantic errors (e.g. multiple columns in the select list of a subquery where only one column is allowed)


### Code completion
The code completion mechanism is context sensitive, meaning that what it offers depends on where you are in the query. **Only objects that can legally be located at the caret position will be offered.** For example, in the `FROM` clause, auto complete will offer tables, views, etc., but will not offer columns. 


When searching for auto complete alternatives, the characters you enter are used as a filter.  
     
![Code completion](http://www.querystorm.com/Images/Docs/autocompletefilter.png)

In the example above, the string *t2* is used as a filter, but **you can skip characters** in the filter which is why writing *t2* will offer `Table2`. For ordering sake it's best to start with the first letter of the object you are searching for, but it's not required. For example, you can write *be2* and it will still offer `Table2` ('Ta**b**l**e2**'). 

Auto-complete will be popping up automatically as you type, but you can turn this off in the settings dialog. You can always invoke auto-complete manually by using the `Ctrl+Space` keyboard shortcut (which you can change in *settings* as well). 

> Tip: when typing columns in the select list, only columns that belong to objects in the FROM clause will be offered, which is why **typing the FROM clause first** is required for column completion in the select list. It's a bit awkward that in SQL, SELECT comes before FROM, as it makes auto-complete much less helpful (without knowing the FROM clause, auto-complete cannot offer meaningful columns in the select list).

### Star expansion
If you invoke auto-complete on a star explicitly (using `Ctrl+Space`), you will be presented with the option to expand the star, i.e. replace it with the list of columns it refers to.

![Star expansion](http://www.querystorm.com/Images/Docs/starexpansion.png)

Names with special characters will be quoted appropriately (depending on the engine being used), and if there are multiple columns with the same name coming from different tables, the columns will be qualified as needed.   

### Column disambiguation
When a column is ambiguous (e.g. between columns from two or more tables), it is underlined as an error. If you invoke auto-complete on such a column (using `Ctrl+Space`) you will be presented with options for disambiguation, as shown in the image below.  

![Column disambiguation](http://www.querystorm.com/Images/Docs/disambiguate.png)

### Tooltips
Hovering the mouse pointer over a symbol will display information about that symbol. For example, here is what it looks like on a star symbol:

![Tooltips](http://www.querystorm.com/Images/Docs/symboltooltip.png)

Tooltips display the type of symbol, what it refers to and (if available) the description. 

### Symbol highlighting

When the caret is positioned on a symbol, all symbols in the document that refer to the same object are highlighted.

![Symbol highlighting](http://www.querystorm.com/Images/Docs/symbolhighlighting.png)  

In the above image, the *p2* table alias is highlighted in two locations where it appears. Table aliases act as  separate copies of the original tables which is why `p1`, and `Person` are not highlighted here. Also, even though there is another symbol named `p2` in the select list, it is not highlighted as it does not refer to the same object.

### Function insights
When calling functions, it's easy to forget their parameters, which is why a good editor (such as this one) will show you this information as you type.

![Function insights](http://www.querystorm.com/Images/Docs/functioninsights.png)  

As you are moving from one parameter to the next, the insights pop-up will highlight the parameter you are currently typing. 

### Bracket matching

When writing complex expressions, it can be very useful to get hints on things like matching brackets:

![Bracket matching](https://www.querystorm.com/Images/Docs/parenmatching.png)

The highlighting works with: 

- quotes
- round, square and curly brackets
- `case`-`end` and `when`-`then` keywords

> **Tip:** To jump the caret between matching segments, you can use the `Ctrl+]` keyboard shortcut. If you also hold `Shift` while you're at it, it will select the entire block between the start and end segments including the segments themselves. 

### Auto formatting
Consistent indentation of the code is very important, it hints at the structure of the query and improves readability. While writing queries it's easy to lose this consistency which is why it's handy if the editor can do it for you.   

In QueryStorm's IDE, you can invoke the auto-format command by pressing `Ctrl+Shift+Enter`. 

Here's an example:

![Auto formatting](https://www.querystorm.com/Images/Docs/autoformat.png)

The autocomplete functionality can be configured to set the capitalization of keywords to Uppercase, Lowercase or Unchanged (in the settings). 

> Note: the auto-formatting functionality will only work if there are no syntax errors in your query

## 2. Object explorer
The object explorer shows database objects in a tree and offers useful context menus for certain items. 

When hovering the mouse pointer over a node, a tooltip is shown with the description of the object (if available).

### Filtering
Searching for objects in the object explorer works much like it does with auto-complete. You can search, and skip letters in the same way. 

One difference is, in the object explorer's filter, you can include multiple criteria separated by a space. A node has to satisfy all the criteria to satisfy the filter. 

Here's an example:

![Object explorer](https://www.querystorm.com/Images/Docs/objectexplorerfilter.png)


### Dragging&dropping nodes into the editor

You can drag&drop a node into the editor (unless it's a category node such as "Tables") and it will drop its name. If the name contains special characters, it will be qualified according to the SQL dialect of the active engine.

If you want to drag and drop all columns of a table, instead of the table name, you can hold `Alt` while dropping the table.

## 3. The results grid
The results grid is mostly self-explanatory, but it has one non obvious feature:

If a cell in the results grid is displaying an Excel address (e.g. `Sheet1!A1` or `A14` or `Sheet1!A1:B10` or `someNamedRange`), **double-clicking that cell will select that range in Excel**!

![Results grid](https://www.querystorm.com/Images/Docs/resultgrid.png) 

Also, if you select multiple cells in the results grid, and some of the cells contain Excel addresses, and you press the `spacebar`, all those addresses will be selected in Excel.   

**This is useful when using the '__address' field that QueryStorm adds to all tables.**

## Other elements of the IDE
The IDE also includes a MessageLog, and a ribbon tab which are quite self-explanatory and do not warrant further explanation here. It also includes a WorkbookQueries pane which is related to automation and is covered there. 