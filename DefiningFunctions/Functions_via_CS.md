# Creating Excel functions

QueryStorm lets you define new Excel functions using **SQL**, **C#** and **VB.NET**. 

These functions can be defined inside:
- A particular **workbook** 
- An **extension package**

Functions that are defined inside a workbook are only available to that workbook. The advantage to this is that they will be immediately available to your users when they open a copy of your workbook (provided they have the QueryStorm runtime).

On the other hand, functions that are defined in an *extension package* can be used in any workbook on the machine that has the package. However, since the functions are not contained inside workbooks themselves, they need to be distributed to your users separately.  

Packages can easily be **published** to allow other users (who have the QueryStorm runtime) to browse, install and use them. For more information on publishing and installing extension packages, [click here](todo). 


## Defining functions with C# #

The process of defining a function is fairly simple:

1. Choose (or create) the project that will host the function
1. Add a new class to the project
1. Write the function in the class
1. Decorate the function with `[ExcelFunction]` attribute
1. Build the project to register the function

Here is a video of the entire process:

[![Defining Excel functions with C#](https://encrypted-tbn0.gstatic.com/images?q=tbn%3AANd9GcTY-5zRYwmgJFGuWvZxc8kSKnSksrbTB5183Q&usqp=CAU)](https://youtu.be/emcyyiVUYSk "Defining Excel functions with C#")

QueryStorm uses the popular **ExcelDNA** library for registering Excel functions. The functions you define can be simple synchronous functions as in the video, but they can also be asynchronous, accept and return tabular values, cache data internally and do other non-trivial things. More information on these topics (and many more) can be found in the following resources:
 - [ExcelDna Google group](https://groups.google.com/g/exceldna/)
 - [ExcelDna on StackOverflow](https://stackoverflow.com/questions/tagged/excel-dna)
 - [ExcelDna wiki on GitHub](https://github.com/Excel-DNA/ExcelDna/wiki)  

## Defining functions via VB.NET

The procedure for VB.NET is the same as for C#, with one distinction - before adding the class file, in the `module.cfg` file, you should change the following line:

```
"Language": "CSharp"
```
...to...

```
"Language": "VisualBasic"
```

That will ensure that the class file template uses VB.NET and that the project builds using the VB.NET compiler. 


## Defining functions via SQL

QueryStorm lets you define Excel functions using SQL, extended with a simple preprocessor syntax.  

The process of defining functions with SQL is outlined [here /*TODO:link*/](http://here).   

!!! Note
When you define a SQL function, QueryStorm generates a C# or VB.NET class (nested below the SQL script file) that defines the function in the same way regular C#/VB.NET functions do, so everything described in this chapter also applies to SQL functions as well.

## Dependencies

Simple functions that do not rely on any shared dependencies can be static and do everything on their own.

However, a functions might need to get hold of a particular service object in order to perform its calculation. That service object might be expensive to create, so we would not want to create a new instance of it in each call. It would be better to have a single instance of the service which we would use in each function call.   

If the function relies on such dependencies, it should not be static. The constructor of the class that owns the function is executed only once (just before the first function call), so any expensive dependencies can be created there.

If multiple functions should share the same dependency, the dependency should be registered centrally, in the IOC container in the `Application` class (inside the *App.cs* file). The constructor can then request it by simply declaring it as a constructor argument (constructor injection). 
 
> If your project does not have an App.cs file, you can create one via the context menu: Add->Application class

For example:

In App.cs
```csharp
public App(IUnityContainer container)
	: base(container)
{
	// register the service as a singleton (ContainerControlledLifetimeManager)
    container.RegisterType<SomeService>(new ContainerControlledLifetimeManager());
}
```

In ExcelFunctions1.cs
```csharp
public class ExcelFunctions1
{
	SomeService someService;
	public ExcelFunctions1(SomeService someService)
	{
		this.someService = someService;
	}

	[ExcelFunction]
	public static int Add(int a, int b)
	{
		return someService.Add(a, b);
	}
}
```

> QueryStorm uses the [Unity container](http://unitycontainer.org/articles/introduction.html) for dependency injection. Some of QueryStorm's own services are available via the IOC container (e.g. `IDialogService`), so you can use those as well. 

## Debugging functions

QueryStorm does not (yet) have a debugger, but there are two methods that help you debug issues: `Log()` and `Debug()`.

Click below to view a video of both debugging methods:

[![Debugging functions](https://encrypted-tbn0.gstatic.com/images?q=tbn%3AANd9GcTY-5zRYwmgJFGuWvZxc8kSKnSksrbTB5183Q&usqp=CAU)](https://youtu.be/zqPGuJoD5DM "Debugging functions")

### The `Log()` method
The simplest way to debug issues is to use the `Log(object obj)` method to print values to the QueryStorm messages pane.

The `Log()` function is contained in the `QueryStorm.Tools.DebugHelpers` class. All class files that QueryStorm generates have a static *using* directive for that class, so you can use the `Log` function anywhere in your code, without qualifying it with the namespace or the class name.  

It's important to note that QueryStorm has two log viewers. One is part of the IDE, and the other is part of the Runtime (launched separately from the ribbon). The output of the `Log()` method will be visible in both places, so Runtime users will be able to see these messages. 

### Attaching a debugger
The Log method is useful, but quite often a proper debugger is needed to track down tricky bugs. QueryStorm compiles code in a way that's friendly to debuggers, so it's fairly easy to debug your code with an external debugger. 

To launch a debugger at a particular location in the source code, use the `Debug()` method. The `Debug()` method is also available anywhere in the code, due to the static using directive `using static QueryStorm.Tools.DebugHelpers;` that's part of all code files QueryStorm generates.

If the local machine has Visual Studio installed, the `Debug()` method will launch Visual Studio, attach it to the process and stop the debugger at the current line. If a debugger is already attached, it will simply stop at the line with the `Debug()` call.

If you do not have Visual Studio installed, you can use the small open source `DNSpy` debugger, attach it to the Excel process, and use the `Debug()` method to stop the debugger at the desired line in the code.     





   

  



