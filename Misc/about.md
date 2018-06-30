# Introduction

Excel might not be a trendy topic these days, but it continues to be used extensively in the business world. People spend a lot of time in Excel and, consequently, there tends to be a lot of data floating around in spreadsheet files. Lack of proper support for tech skills in Excel, though, is a constant source of frustration for many tech users and is costly for them, their employers, business users and clients. The purpose of QueryStorm is to greatly reduce this problem by adding easy and convenient support for SQL and C# into Excel, as well as better database connectivity.

## What can you use QueryStorm for?

### Processing incoming Excel data
Data in Excel is usually tabular, and quite often relational, but support for this kind of data in Excel is quite poor. VBA, which comes with Excel, is not well suited for working with tabular data. It works with cells and ranges and does not offer set based and relational operations. QueryStorm fixes this shortcoming of Excel by including support for SQL and C#, which are both well suited for working with sets of data.

![Screenshot](https://i.imgur.com/TyDNgMP.png)

Since business users tend to deal with data in spreadsheets, and tech users deal with data in databases, there's often a need to move data between the two. There are many tools that do this, but they tend to be clumsy and involve a lot of manual preparation. QueryStorm makes moving data between Excel and databases much easier. As soon as you connect, Excel tables are visible to the database (as temp tables). From there you can import them into permanent tables, or query them alongside permanent tables. Query results are easily returned into Excel, so moving data the other way is also a breeze.

![Connected to DB](https://i.imgur.com/3Fqom4n.png)

### Reporting
Spreadsheets are a popular way of presenting data, due to having support for sorting, filtering, graphs and pivot tables. The data usually comes from databases but can come from anywhere really. Some times the reports are static. Other times they need to be interactive, e.g. the user changes a parameter in a cell which causes a query (or script) to re-execute and update the data in a table, which in turn updates a chart.

![Example QueryStorm report](https://www.querystorm.com/downloads/images/exappgif.gif "Example QueryStorm report")

You probably know of at least a few ways to get data into Excel (e.g. PowerQuery aka Get&Transform). Tools for this are usually GUI based and, while more-or-less easy to use, they tend not to be entirely flexible. QueryStorm on the other hand takes a very tech-centered approach. When you connect to a database, you interact with it using SQL and a full-featured editor. The data is usually processed in the database and only the results are ever returned into Excel. In addition, you can load data from other sources into Excel using C# and various .NET libraries. There are no extra layers of abstraction, no new languages to learn (e.g. DAX, M) and no tabular model to populate and refresh. 

There are certainly many situations where GUI based tools are more appropriate, but if you're fairly technical, there's also a good place for QueryStorm in your reporting toolkit.

### Prototyping applications
Excel offers a lot out the box: UI, data storage and visualization. If it came with a good programming language (e.g. C# instead of VBA), tech users could use it effectively for rapidly building prototypes and small applications. 

![C# automation in Excel](https://i.imgur.com/HIPI5Gu.png)

With C# you can load data from anywhere, and store it in the Excel file. As soon as the data is written into an Excel table, you immediately get strongly typed access to it (as if you had an ORM) so you can use LINQ to work with it. You don't need to spend time setting up a database and making sure it's accessible.

Excel also takes care of most of the UI work. You can set up graphs and user interactions (buttons, slicers, cell triggers...) much faster, compared to building a dedicated application. 

If you need to generate reports, you can easily define a template in a sheet, populate it with dynamic data and print it out as a pdf file. 
 
### Tabular data scratchpad
Data is all over the place and, in general, the more techy you are, the more you interact with it. Suppose you need data for a machine learning model. You load bits and pieces of data from multiple sources: APIs, csv files, xlsx files, databases... Usually data collected this way needs to be cleaned and molded before can put it to good use. 

With QueryStorm, we can easily access data in databases as well as data from elsewhere (via C#). As soon as the data is in Excel, you immediately get strongly typed access to it so you can work with it via code. You also get the benefits of a spreadsheet - you can work with it by hand, add formulas, visualize it (conditional formatting, graphs) and casually explore it.

This makes fetching, cleaning and molding data much easier. And since it's all in scripts and queries, it's repeatable. 

### Teaching SQL
Excel is quite visual and pretty understandable to most people. This makes QueryStorm an excellent tool for teaching and learning SQL, since you get to view the data and the effects of queries it real time. That's why QueryStorm is being used in several educational institutions and companies to teach and learn SQL. If you're interested, check out the [SQL crash course for Excel users](https://www.udemy.com/sql-fundamentals-for-excel-users/learn/v4/overview "SQL fundamentals for excel users"), it uses QueryStorm. 

## QueryStorm works on Excel for Windows only
QueryStorm is a plugin for the desktop version of Excel for Windows. Due to technical limitations it does not work on a Mac. There is no web version either, so it's not available in Google Sheets or Excel online. That does not mean you can't work with data in e.g. Google Sheets. If you have Google Sync, you can open a file that's on your Google drive locally, and every time you save your work, it will be automatically synced online, where others can see your changes.


## Free for non-commercial use
QueryStorm is free for non-commercial use. You can use it freely for your own needs. The only limitation currently is connectivity to external databases. If you are using it commercially though, please buy a license.

## QueryStorm Runtime license
Tech users are usually going to be the ones writing code and embedding it into Excel files. If those files are sent to business users, they also need to have QueryStorm installed for those scripts and queries to be able to execute. The business users however don't need to write code themselves so they do not need a full license for QueryStorm. Runtime licenses are available for purchase, and are considerably cheaper than the full ones. They enable embedded code to run, but they don't see the code editing features:

![Ribbon for QueryStorm runtime](https://www.querystorm.com/downloads/images/runtime_ribbon.png)

## Who's behind QueryStorm

In alphabetic order:

- [Antonio Nakic-Alfirevic](https://www.linkedin.com/in/anakic), on the keyboard (developer, founder).
- [Benjamin Van Dam](https://www.linkedin.com/in/benjamin-van-dam-992237b9/), on the mic (word spreader).
- [Boško Kustura](https://www.linkedin.com/in/bo%C5%A1ko-kustura-9b16102/), on the drums (operations).


