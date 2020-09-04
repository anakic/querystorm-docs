# The data context

A data context exposes tables, variables and events to your scripts and components.

When you create a script, a new instance of the datacontext is automatically created and passed to the script. The script gets its data from the data context, and it can only see tables and variables that the data context provides.

In each project there is exactly one data context class. If you do not define a data context class, a default one is used. 

You may want to create your own data context class to customize the data that will be available to your scripts and components.

## The `WorkbookDataContext`

For workbook projects that do not define a data context, a `WorkbookDataContext` is used by default.

A `WorkbookDataContext` instance exposes data from a particular workbook in the following way:

- Tables = Excel tables
- Variables = Single-cell named ranges
- Events = workbook events (e.g. button click)

## Custom data context

Overriding the data context allows you to:

- Add additional tables
- Change column data types
- Add relations between tables

To define a data context class, right click on the project node and click "Add"=>"DataContext class".

![Add data context](../images/add_datacontext_menu.png)

### Adding tables

To include additional tables in the data context, add them to the `Tables` collection inside the `Initialize` method, as shown below:

``` csharp
protected override void Initialize(ContextSchema schema)
{
    // an example table
    var myFilesTbl = new TabularValue(
        System.IO.Directory.GetFiles(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        )
        , "myFiles"
    );
    // add it to the context
    Tables.Add(myFilesTbl);
    
    base.Initialize(schema);
}
```

The `DataContext.Tables` collection contains objects that derive from the abstract class `Tabular`. You can choose from the existing implementations of `Tabular`, or create your own. For example, `TabularValue` is a versatile implementation that accepts any object (usually a collection) and exposes its values as a table.

### Changing column data types

Excel does not enforce consistent column types, so the data context has to guess at their types based on their content. For example, suppose a column contains the numbers `1,2` and `3`. The auto-detect mechanism sees that the column contains numbers, but it does not know if the column is allowed to contain nulls or floating point numbers so it selects `double?` as the data type. 

To specify the column type explicitly, use the `schema` object in the `Initialize` method as shown below:

``` csharp
protected override void Initialize(ContextSchema schema)
{
	schema.ConfigureTable("Person")
		.ConfigureColumn<System.Int32>("Id")
        .ConfigureColumn<System.String>("Name")
		.ConfigureColumn<System.DateTime>("DateOfBirth");
		
	base.Initialize(schema);
}
```

### Table relations

Excel also doesn't offer a way to define relationships between tables, but you can define those relationships inside a data context class. 

Each time the data context file is saved, it is automatically compiled and loaded, and strongly typed classes are generated based on the tables it offers. Specifying relationships between tables, ensures that you get **strongly typed navigation properties** on those classes.

This does not impact SQL scripts but it does impact component code as well as C# scripts. 

To add table relations, use the `schema` object in the `Initialize` method as shown below:

``` csharp
protected override void Initialize(ContextSchema schema)
{
	schema.ConfigureTable("Departments")
		.ConfigureColumn<System.Int32>("Id")
		.ConfigureColumn<System.String>("Name");

	schema.ConfigureTable("Employees")
		.ConfigureColumn<System.Int32>("Id")
		.ConfigureColumn<System.String>("FirstName")
		.ConfigureColumn<System.Int32>("DptId")
		// adding a relation from Employees.DptId to Departments.Id
		.AddRelation("DptId", To.One, "Departments", "Id", "MyDepartment");
	
	base.Initialize(schema);
}
```

Once the project is compiled, C# scripts will be able to use the new property `MyDepartment` on the employee objects.

![CSScript relationship navigation](../images/relationship_navigation_example.png)

> Tip: for performance reasons (to avoid type conversions), it's slightly better to leave Id columns as `double` instead of `int`. 

### Demo

Click below for a video demonstration of the contepts discussed in this segment:

[![Customizing the data context](../Images/video.jpg)](https://youtu.be/q9tv1h5j3o4 "Customizing the data context")
