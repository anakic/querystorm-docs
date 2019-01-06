# CSV files
External CSV files can be consumed as data sources from both C# and SQLite.

In the following example, we're reading a list of items from a CSV file named "run history.csv" using C#:

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
public static CsvCommittableCollection<RunHistory> ReadCsv_RunHistory()
{
    return Csv<RunHistory>(@"C:\Users\anton\AppData\Roaming\Everything\Run History.csv", ",");
}

```

!!! Note
	The properties of the `RunHistory` class above are mapped to headers in the CSV file. If the property and the header name don't match (e.g. if the header has spaces), the header name can be specified via the `[Name]` attribute on the property, as demonstrated above.

We can now use this function from C# to read the data in the CSV file. If we make sure to decorate the function with `[SQLFuncTabular]` and embed it into the workbook, we'll also be able to use it from SQLite.

The object that the CSV method returns supports committing changes via the `Commit` method. 

## Generating the code

A quick query can be used to generate this code automatically:

![Generate CSV accessor](../images/qqcsv.png)

Clicking "Execute transformation" will generate the C# code and embedded it into the workbook automatically. 

![Generated CSV code](../images/csv_code.png)

The quick query does not detect types in the csv file. Rather, it just assumes all columns are of type string. In the example above, it's reasonable to assume that the column `RunCount` will contain only integers. In this case, it's safe to change the property type to `int`.

## Accessing the CSV file from SQLite
Since the C# function has the `[SQLFuncTabular]` attribute applied, if we embed this script into the workbook, we will be able to use the generated `ReadCsv_RunHistory()` function as a table-valued function from SQLite. 

![Reading a CSV file from SQLite](../images/csv_from_sql.png "Reading a CSV file from SQLite")

## Decimal separator & date format

Suppose your CSV file uses the Polish style of dates and numbers (e.g. coma as decimal separator) and uses a semi-colon as a column separator. You can specify the column separator and the culture information like so:

```csharp
return Csv<qs_csv>(@"C:\Users\anton\Downloads\qs_csv.csv", ";", new System.Globalization.CultureInfo("pl"));

```
