# Managing packages

QueryStorm comes with NuGet support, which serves two purposes:

- Installing NuGet packages into projects
- Sharing QueryStorm extension packages

## NuGet packages

The package manager is used for managing NuGet packages in QueryStorm projects (viewing, installing, uninstalling, updating).

![PackageManager](../../Images/PackageManager.png)

Installing a package into your project does the following:

- Adds an entry in the `Dependencies` section of the `module.config` file
- Downloads the package
- Unpacks the contents of the package into the project
  
All dll files contained inside the package are unpacked into the `lib` folder and are automatically referenced.

Since the package contents get stored inside the project, installing NuGet packages into a workbook project will increase the size of the workbook file. Most NuGet packages are fairly small, however, so the increase in size is unlikely to be an issue. The advantage to this is that workbook applications are self-contained and will run on any machine that has the QueryStorm runtime.
