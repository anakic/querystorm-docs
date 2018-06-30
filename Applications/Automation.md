# Automation
As mentioned in the introduction, QueryStorm consists of four primary components:

- the IDE
- the built in database engine (augmented SQLite)
- external database connectors
- **the automation runtime** 

The automation runtime enables you to set up your workbook so that your queries are connecting, executing and writing results in the background as a response to any triggers that you set up. 

You can set up most automation scenarios without the use of VBA, but QueryStorm does expose a VBA API which gives the most flexibility with regards to automation. For information on automating QueryStorm through VBA, take a look at the *VBA API* section.

## How does it work?
### Embedded queries
To enable automation, QueryStorm needs to **embed queries** into a hidden area of the workbook that is designed for such custom data. You can see all queries that are embedded in the workbook if you take a look in the "*Workbook queries*" pane in QueryStorm. 

![Data in a sheet](https://www.querystorm.com/Images/Docs/workbookqueries.png)

You can create embedded queries yourself by clicking the "+" icon in "*Workbook queries*" pane. Embedded queries are also created automatically by the table writer when **SetupRefresh** is set to `Manual` or `Automatic`.

> Note: there is a **naming convention** for embedded queries which serves to associate queries and tables. For example, if you have a table called `Table1` and your embedded query is called `Table1.TestQuery`, the query will show up in the context menu of Table1 in the "Related queries" submenu. The naming convention consists of starting the query name with '*[nameOfTable].*' followed by any name you choose. The only purpose of this naming convention is to make the query available in the context menu of the table.

#### *Anatomy* of an embedded query
An embedded query will usually consist of these four elements:

1. **the SQL** query that gets the results
- the **@-directive** that scripts writing the results to the workbook
- the **engine configuration** that will be used to run the query
- the list of **triggers** that will trigger the execution of the query

The four elements are shown in the two images below.

![Data in a sheet](https://www.querystorm.com/Images/Docs/embeddedquery_anatomy_1_2.png)

The SQL and the @-directive (elements 1 and 2) are the actual content of the embedded query, while the engine configuration (3) and triggers (4) are the query's metadata which you can see and configure by clicking the automation (little robot) icon.

![Data in a sheet](https://www.querystorm.com/Images/Docs/embeddedquery_anatomy_3_4.png)

The embedded query can use the Live mode engine or any of the External mode engines. 

When an embedded query is triggered (or manually executed from the context menu of an associated table), it will create and open its own connection (which potentially involves copying data to temp tables), execute the query, write any results if needed, and finally close its connection. This is all done in the background and does not freeze Excel's user interface, nor does it interfere with QueryStorm's UI and any operations you might be doing with it.   

### Triggers
There are several types of triggers you can use to automate your embedded queries:

- Workbook: Loaded
- Table: LostFocus, Changed
- Range: LostFocus, Changed
- Timer: (specify the interval in milliseconds)
- Button: Click (button must be the ActiveX kind)

You can specify the triggers for an embedded query in the automation dialog (which opens when you click the little robot icon), but there are also handy context menus that are available for this purposes. The context menus will appear for the following Excel objects: 

- Tables
- Cells
- Buttons (the ActiveX kind)

This makes it easy to specify that the click of a button or the changing of a table/range should trigger the execution of an embedded query. 

Here's what the context menu looks like on a button:
 
![Data in a sheet](https://www.querystorm.com/Images/Docs/buttontriggers.png)



----------
 
### So what does all this do to your workbook?
Embedding queries doesn't do anything funky with your workbook, it just writes a bit of text into a hidden area (xls and xlsx files have this area, and it's called *CustomXMLParts*). The workbook will show up perfectly fine even on machines that don't have QueryStorm installed. If there's no QueryStorm, then automation will not run, but the workbook will be perfectly fine and will work like a normal Excel workbook.