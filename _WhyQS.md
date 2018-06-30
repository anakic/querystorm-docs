# Why QueryStorm?

It might seem like we have an application for everything these days, but it's hardly the case. We have custom and adhoc requirements all the time, and we often use Excel to handle them. If our accounting application doesn't do everything we need it to, chances are we have a spreadsheet or two on the side that we use for the purpose. As a result, a sizable flow of data in today's organizations relies on Excel. Whether we like it or not, spreadsheets are one of the primary ways business users consume and exchange data, and in order to maintain a healthy flow of data throughout an organization, this needs to be taken into account. 

### Processing data
Different applications often can't talk to each other, but many of them can export data into Excel. As a result of this, a lot of data comes into, goes out of and circulates inside of our companies in the form of Excel files. In order to make use of this data, we usually need to clean and transform it, extract conclusions from it and/or move it into databases. This data is usually tabular and often relational so being able to use SQL and/or C#(linq) is quite convenient for processing it. 

An [annual survey](http://www.oreilly.com/data/free/files/2016-data-science-salary-survey.pdf "2016 Data Science Salary Survey.pdf") on salaries in the field of data science done by O'Reilly Media, found that Excel and SQL are the two most popular tools that data scientists use in their work, followed by R and Python. While not conclusive, this suggests a large amount of data is exchanged, stored and analyzed in Excel and databases. QueryStorm helps with this by making use of the user's existing knowledge of SQL and/or C# to enhance the way they work with data in Excel, as well as making it much easier to move data in either direction between Excel and databases.

### Building Excel applications  
Excel isn't just about storing and exchanging data. The primary thing that made spreadsheets popular is automatic calculation capabilities in the form of formulas. This was introduced by Excel's ancestor [VisiCalc](https://en.wikipedia.org/wiki/VisiCalc "VisiCalc wiki") and made it (and Apple 2, which it ran on) immensely popular. 

People could use formulas to build all sorts of applications, for example workbooks for tax calculations and book keeping. With the introduction of VBA, programming was introduced to spreadsheets in a more serious way.

In this sense, a workbook is an application running on Excel as the platform.  

Today, there are over 1 billion installations of Excel worldwide, half of which are in businesses. This makes Excel a quite the platform. You can send a workbook to pretty much anyone, and be confident they'll be able to open and run it. 

This is all well and good in theory, but in reality the applications we build in Excel are quite fragile resulting in costly errors and some [high profile debacles](http://fortune.com/2013/04/17/damn-excel-how-the-most-important-software-application-of-all-time-is-ruining-the-world/ "Excel debacles"). It's quite easy to reference the wrong cell in a formula or mess something up in VBA code, especially in long lived workbooks maintained by many people.

While software aren't blameless, having better tools can improve the outcomes as well as make work easier on the people doing it. QueryStorm helps by introducing support for C# into Excel as a data processing and automation language. Compared to VBA it is much better equipped at dealing with tabular data (LINQ), is a more modern and expressive language (VBA last updated in 2010), has better tooling and a much wider and more modern ecosystem of libraries available to it. 


## About QueryStorm

Conceptually, QueryStorm consists of the following parts:

- The IDE 
- The built-in SQLite engine
- The built-in C# engine 
- Connectors to external database engines 
- The automation runtime 

Emphasis has been placed on keeping the user interface simple and the learning curve small while still adding a huge amount of power to the user. 

SQL and C# are at the heart of what QueryStorm does, so to use it effectively, being comfortable with either is pretty much required. If you aren't though, you still might find it useful, though. QueryStorm is an awesome playground for practicing SQL queries, so it can and is being used to **[teach and learn SQL](ubaciti udemy link)**.


### QueryStorm works with tables
 
The most important thing to know right off the bat is that **QueryStorm works with Excel tables** as opposed to sheets and ranges. The reason is that, with tables, the data and headers are clearly marked, and you can have whatever you like around them. You can also have multiple tables in the same sheet and QueryStorm will happily work with them without any issues.

|Just a range| Ctrl+T  |Now it's a proper table|
|---|---|---|
| ![Data in a sheet](http://querystorm.com/Images/Guide/range.png)  | ![the Insert table button](http://querystorm.com/Images/Guide/inserttable.png)  |  ![data marked as a table](http://querystorm.com/Images/Guide/table.png) |

    

To mark a range as a table, select the range (or just one or more cells in it) and click *Insert->Table* in the ribbon (or press `Ctrl`+`T`). You can change the table's name in the “Design” tab in the ribbon. 

That's all the preparation we need, now we can get querying.

----------
###### Last updated: 11/16/2017 5:01:57 PM