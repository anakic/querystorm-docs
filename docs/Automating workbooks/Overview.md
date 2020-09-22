# Automating workbooks

Excel files offer data storage capabilities and a flexible user interface for data entry and visualization. However, Excel files aren't just static documents containing data; they can also have behavior, in effect, making the workbook a hybrid between a document and an application.

Out of the box, Microsoft Excel allows automating workbooks using VBA. While this works, VBA is quite old and hasn't kept up with the new programming paradigms and ecosystems. QueryStorm offers the ability to use modern programming languages and ecosystems when writing code to automate workbooks.

This allows using C# or VB.NET code in workbooks, instead of VBA. It also allows using NuGet packages and existing .NET dlls inside Excel workbooks, making it easy to leverage your existing proprietary and 3rd party code in workbook applications.

## Model-binding

A simple way to automate a workbook is to simply access Excel via it's COM API. While easy to get started with, writing code for interacting with the Excel API can get quite tedious. That's why QueryStorm offers a model-binding approach to automation, which minimizes the amount of Excel interaction code, and lets you focus on the logic of your application.

## SQL-based automation

QueryStorm also allows setting up automation from SQL scripts. These scripts typically contain SQL code mixed in with a preprocessor syntax. The SQL code fetches or updates data in a database. A preprocessor syntax is added to allow sending query results into the workbook, and to specify which events will trigger the execution of the script.

In QueryStorm, all automation is done using C# or VB.NET code. Under the hood, though, SQL scripts generate this code for you. This lets you focus on SQL, rather than on writing component code. It is also very useful for the many SQL users who are not familiar with .NET programming languages.

## Application lifetime

QueryStorm workbook applications are started (by the QueryStorm runtime) as soon as the project is built, or when the workbook is opened. If we make changes and rebuild the project, the old instance of the application is unloaded and the new one is started.