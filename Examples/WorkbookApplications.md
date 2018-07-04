## Building workbook applications
With C# we can build workbooks with rich behavior, making use of the .NET ecosystem. Coupled with the benefits of a spreadsheet (storage, UI, graphs, self contained, easy to share), this makes building prototypes an adhoc applications really easy and fast. In this sense, C# is a replacement for VBA.

> Still missing is a debugger, which is planned later this year.

### Example #1: Exchange rates visualizer
Suppose you want to see how the US dollar has been doing in the past 36 months. We can use C# to load exchange rates data from a web API and use Excel to store and visualize this data.

The script below generates 36 dates, one for each month back in history, and use the fixer API to get exchange rates for each data. Once the data is ready, it is dumped into an Excel table so it can be visualized.

``` C#
using Newtonsoft.Json.Linq;

int numberOfMonths = 36; //get data for the last 36 months
string baseCurrency = Range("baseCurr").ValueString(); //rerad base currency from cell

//fetch data
var data = Enumerable
    .Range(1, numberOfMonths) //generate integers
    .Select(i => DateTime.Today.AddMonths(-i)) //convert to dates
    .SelectMany(dt =>
    {
 		//output: date, currency, rate
        return GetExchangeData(dt, baseCurrency)
            .Select(ed => new
            {
                Date = dt,
                Currency = ed.Currency,
                Rate = ed.Exchange
            });
    });

//push results into Excel table
data.IntoTable(nameof(rates));

//gets the exchange rates for a specified base currency on the specified date
private static IEnumerable<Rate> GetExchangeData(DateTime date, string baseCurrency = "EUR")
{
    var response = new WebClient()
        .DownloadString($"https://api.fixer.io/{date.ToString("yyyy-MM-dd")}?base={baseCurrency}");

    string ratesStr = JObject.Parse(response).Property("rates").Value.ToString();

    return JsonConvert
        .DeserializeObject<IDictionary<string, decimal>>(ratesStr)
        .Select(kvp => new Rate() { Currency = kvp.Key, Exchange = kvp.Value });
}

public class Rate
{
    public string Currency { get; set; }
    public decimal Exchange { get; set; }
}
```

And here is the result:

![]()

### Example #2: Simple RSS reader application
In the following video, I demonstrate building an RSS reader application in Excel:

[![Screencast](https://i.imgur.com/KBTKw5u.png)](https://vimeo.com/241888977 "RSS Reader demo on Vimeo")


### Example #3: Loading data from a database
QueryStorm comes with its own database engine (SQLite) and support for connecting to external databases (SQL Server, MySql, Postgres, Access, ODBC). This makes it quite easy to fetch data from databases.

Once prepared, queries can easily be called from C#, and they return data as `IEnumerable<dynamic>` objects.  

Here's an example...

**Step 1:** Connect to SQL Server: 

```
Server=mssql6.mojsite.com,1555; Database=thingieq_AdventureWorks2014; User Id=thingieq_abc; Password=D8R7VXLcSzss
```

**Step 2:** Prepare and the SQL query, and embed it into the workbook e.g. as *"Products query"*:

``` SQL
SELECT * FROM Production.Product
```
**Step 3:** Call the query from C#

```C#
//execute query to fetch data from the database
var result = await Query("Products query").RunAsync();

//items collection is of type IEnumerable<dynamic>
var data = result.Items;

//Output data into Excel table
data.IntoTable("Parts", "B2")
```

Here's the result:
![running db query](https://i.imgur.com/mG7Hrrq.png)

If we like, we can embed the C# script into the workbook and add a job that automates it. 