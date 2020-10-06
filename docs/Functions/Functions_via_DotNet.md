# Defining functions with .NET

The process of defining a function with C# or VB.NET is simply a matter of writing the function and decorating it with the `ExcelFunction` attribute.

```csharp
public class MyFunctions
{
    [ExcelFunction]
    public static int Add(int a, int b)
    {
        return a + b;
    }
}
```

For a video demonstration click below:

![YOUTUBE](emcyyiVUYSk)

## Loading the new function

Once the project that contains the function is built (compiled), the runtime will automatically load it and make the function available in Excel. Depending on if the function is in a workbook or an extension project, the function will be available in the defining workbook or all workbooks.

## ExcelDNA

QueryStorm uses the popular **ExcelDNA** library for registering Excel functions. The functions you define can be simple synchronous functions as shown here, but they can also be asynchronous, accept and return tabular values, cache data internally, and do other non-trivial things.

More information on these topics, and many more, can be found in the following resources:

- [ExcelDna Google group](https://groups.google.com/g/exceldna/)
- [ExcelDna on StackOverflow](https://stackoverflow.com/questions/tagged/excel-dna)
- [ExcelDna wiki on GitHub](https://github.com/Excel-DNA/ExcelDna/wiki)

## C# or VB.NET

When creating the project, you can choose which language to use:

![New project dialog](../Images/new_project_dialog.png)

The language setting is stored in the `module.config` file.

```json
"Language": "CSharp"
```

...or...

```json
"Language": "VisualBasic"
```

The language setting determines the compiler and class templates that the project will use. If you choose VB.NET, the code will be compiled using the VB.NET Roslyn compiler, and the scaffolded function files will look something like this:

```vb
Public Module MyFunctions1
        <ExcelFunction>
        Public Function Add(val1 As Int32, val2 As Int32) as int32
            return val1 + val2
        End Function
End Module
```

## Dependencies

Simple functions that do not rely on any shared dependencies can be static and do everything on their own.

However, a function might need to get hold of a particular service (e.g. an API object) in order to perform its calculation. That service might be expensive to create, so we would not want to create a new instance each time the function is evaluated. It would be better to have a single instance of the service which would be reused in each function call.

If the function relies on such dependencies, it should not be static. The constructor of the class that owns the function is executed only once (just before the first function call), so any expensive dependencies can be created there.

However, if multiple functions should share the same dependency, the dependency should be registered centrally, in the IOC container in the `App` class (inside the App.cs file). The constructor of the function's class can then request the service by simply declaring it as a constructor argument (constructor injection).

For example, we can register the service in App.cs:

```csharp
public App(IUnityContainer container)
    : base(container)
{
    // register the service as a singleton
    container.RegisterType<SomeService>(new ContainerControlledLifetimeManager());
}
```

We can then use the service in our function:

```csharp
public class ExcelFunctions1
{
    SomeService someService;
    // request the service by adding a ctor argument
    public ExcelFunctions1(SomeService someService)
    {
        this.someService = someService;
    }

    [ExcelFunction]
    public int Add(int a, int b)
    {
        // use the service to perform the calculation
        return someService.Add(a, b);
    }
}
```

> QueryStorm uses the [Unity container](http://unitycontainer.org/articles/introduction.html) for dependency injection.

## Debugging functions

QueryStorm does not currently have a built-in debugger, but there are two static methods that can help with debugging: `Log()` and `Debug()`.

![YOUTUBE](zqPGuJoD5DM)

### The `Log()` method

The simplest way to debug issues is to use the `Log(object obj)` method to print values to the messages pane.

The `Log()` method is contained in the `QueryStorm.Core.DebugHelpers` class. All class files that QueryStorm generates have a static *using* directive for that class, so you can use the `Log` method anywhere in your code, without qualifying it with the namespace or the class name.  

It's important to note that QueryStorm has two log viewers. One is part of the IDE, and the other is part of the Runtime (launched separately from the ribbon). The output of the `Log()` method will be visible in both places, so Runtime users will be able to see these messages.

### Attaching a debugger

The `Log()` method is useful, but quite often a proper debugger is needed to track down tricky bugs. QueryStorm compiles code in a debugger-friendly way, so it's fairly easy to debug your code with an external debugger.

To launch a debugger at a particular location in the source code, use the `Debug()` method. The `Debug()` method is also available anywhere in the code without prefixing it with the namespace or the class name, due to the `using` directive that's part of all code files generated by QueryStorm.

If the local machine has Visual Studio installed, the `Debug()` method will launch Visual Studio, attach it to the process and stop the debugger at the current line. If a debugger is already attached, it will simply stop at the line with the `Debug()` call.

If you do not have Visual Studio installed, you can use the small open-source [DNSpy](https://github.com/0xd4d/dnSpy) debugger, attach it to the Excel process, and use the `Debug()` method to stop the debugger at the desired line in the code.
