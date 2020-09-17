# The App class

In QueryStorm, you can define projects inside your workbook (workbook projects), or in a local folder (extension projects). Workbook projects serve to automate the workbook and define functions that are only available inside that particular workbook. Extension projects serve to define functions that should be usable in any Excel workbook.

Each project is a small **application** that the QueryStorm runtime runs and manages. The `App` class is the entry point to this application.

![The app class node](../../../Images/AppClass.png)

> If your project does not define an `App` class explicitly, a default `Application` instance is used implicitly.

When loading the module, the Runtime will create a single instance of the `App` class. It is the job the `App` class to initialize the application i.e. to register services, create and initialize the [data context](todo) and [components](todo).

The constructor of the App class is the first method that runs when the QueryStorm runtime loads the application. In it, you can do any initialization logic that your applications needs. Applications support dependency injection, and you should use this constructor to register any services that your application needs.

## Dependency Injection

The App class can request certain services from the QueryStorm Runtime. These services are injected as constructor arguments using a Unity IOC container.

For example, if your application needs to interact with the workbook, it can request the `IWorkbookAccessor` service, by adding it to the list of constructor arguments, like so:

```csharp
public App(IUnityContainer container, IWorkbookAccessor workbookAccessor)
			: base(container)
{
    var workbook = workbookAccessor.Workbook;
    // todo: do something with the Workbook object
}
```

The following services are registered with the container out of the box (by the QueryStorm Runtime):

- `IExcelAccessor`: allows access to the current Excel Application instance.
- `IWorkbookAccessor`: Allows access to the workbook that contains the application (defined only for Workbook applications).
- `IDialogService`: allows showing dialog messages to the user, using the QueryStorm style of dialog windows and ensuring that Excel is set as the owner of the dialog windows.
- `CredentialsVault`: gives access to the current user's encrypted connection credentials. This is primarily used by SQL scripts and related classes.

If any other services are required by your application, you should register them with the container in the constructor of the `App` class. Since the container instance belongs to your application, the changes that you make to it will not affect anything outside of your application.

## Workbook applications

QueryStorm lets you automate your workbooks with C# or VB.NET code. This offers many advantages compared to using VBA, most notably better language features and access to the .NET ecosystem via NuGet packages.

In addition, QueryStorm offers strongly typed access to Excel tables, and a model-binding API that lets you concentrate on your business logic instead of on communicating with Excel.

## Extension applications

Extension applications are defined inside projects at the machine level. 

The only purpose of extension projects (for now) is to define a set of custom Excel functions that you can use in any workbook, and that you can share as a package to other (QueryStorm runtime) users. 

Even though they only hold functions, they are loaded by the QueryStorm runtime in the exact same way as Workbook applications. They have an `App` class that they can use to register dependencies that the functions can share.
If it's a commercial package, we can also check for a valid license to make sure the user is licensed to use the package.

## Application lifetime
The runtime is responsible for loading applications. Each applications is loaded into a separate `AppDomain`, ensuring a level of separation between different user applications. 

Applications can be defined at the workbook level (Workbook applications) and at the machine level (Extension applications).

A Workbook application is loaded as soon as the workbook (that contains it) is opened, and is unloaded as soon as the workbook is closed.

Extension applications are loaded when Excel opens, and are unloaded when Excel closes. They are not affected by workbooks being opened or closed. Additionally, when an extension package is downloaded via the Extensions manager, it is immediately loaded. When it is removed via the Extensions manager it is immediately unloaded.

Both Workbook and Extension applications are reloaded each time a new version of the appliciation is built by the IDE. When the IDE successfully builds a project, it sends a message to the Runtime to inform it that a new module has been prepared, and the Runtime immediately attempts to load it, after terminating any previously running instances of the same application.

When unloading, the AppDomain that held the App instance is destroyed, so you don't need to worry about cleaning.
