# CSV files
CSV files can be consumed as data sources from both the C# and SQLite engine without having to explicitly load the file into Excel.

Here's an example of reading a list of items from a CSV file using C#:

```csharp
public class RunHistory
{
    public string Filename { get; set; }
    [CsvHelper.Configuration.Attributes.Name("Run Count")]
    public int RunCount { get; set; }
    [CsvHelper.Configuration.Attributes.Name("Last Run Date")]
    public string LastRunDate { get; set; }
}

[SQLFuncTabular]
public static IEnumerable<RunHistory> ReadCsv_RunHistory()
{
    return Csv<RunHistory>(@"C:\Users\anton\AppData\Roaming\Everything\Run History.csv", ",");
}

```

The properties of the class are mapped to headers in the CSV file. If the property and the header name don't match (e.g. if the header has spaces), the header name is specified via the [Name] attribute on the property.

We can now use this function from C# to read the data from the CSV file. If we make sure to decorate the function with `[SQLFuncTabular]` and embed it into the workbook, we'll also be able to use it from SQLite, as demonstrated below.

## Generating the code

A quick query can be used to generate this code automatically:

![Generate CSV accessor](https://www.querystorm.com/Downloads/Images/qqcsv.png)

Clicking "Execute transformation" will generate the script and embedded it into the workbook automatically. The properties in the generated class are always of type string. You can, of course, modify the column (property) types if you have columns that are guaranteed to be integers, dates or other data types.

## Accessing the CSV file from SQLite
Since the C# function has the `[SQLFuncTabular]` attribute applied, if we embed this script into the workbook, we will be able to use the generated `ReadCsv_RunHistory()` function as a table-valued function from SQLite. 

![Generate CSV accessor](https://www.querystorm.com/Downloads/Images/csvsql.png)

## Decimal separator & date format

Suppose you CSV file uses the Polish style of dates and numbers (e.g. coma as decimal separator) and uses a semi-colon as a column separator. 

You can specify the column separator and the culture information like so:

```csharp
return Csv<qs_csv>(@"C:\Users\anton\Downloads\qs_csv.csv", ";", new System.Globalization.CultureInfo("pl"));

```
