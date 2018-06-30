# Custom table-valued functions

Suppose you're comfortable with C# but your colleagues only know SQL. In QueryStorm, you can write C# table-valued functions which you and your colleagues can call from SQLite. You can expose all sorts of data to SQLite this way e.g. data from web services, REST APIs, local files, active directory, external databases... you name it!

> You can download the demo workbook from [here](../demofiles/tvf_currencies.xlsx).

## Defining the function

Let's suppose we need some data about exchange rates that we want to correlate with data in an Excel table. For a start, let's conn	ect via C# and paste in the following function into the editor: 

```csharp
using Newtonsoft.Json.Linq;

///<summary>Gets the exchange rates for the specified currency on the specified date.</summary>
///<param name="date">The date for which to get the currency rates for</param>
[SQLFuncTabular]
private static IEnumerable<Rate> GetExchangeData(DateTime date, string baseCurrency = "EUR", string symbols = null)
{
    var requestUri = $"https://exchangeratesapi.io/api/{date.ToString("yyyy-MM-dd")}?base={baseCurrency}&symbols={symbols.Replace(" ","")}";
    var response = new WebClient().DownloadString(requestUri);

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

!!! Note
	The `Newtosoft.Json` library is included in QueryStorm, and we can use it right out of the box.

Our `GetExchangeData()` function takes in the desired date, a base currency, and the desired currencies to return. It hits a web API to load the data and parses the JSON response into a collection of `Rate` objects. 

We can of course call this function from C#, but we can also expose it to SQLite. To do so, we need to to do the following:

1. make sure it successfully compiles (has no compile-time errors)
1. decorate it with the `[SQLFuncTabular]` attribute 
1. embed it into the workbook

Once that's done, it will be available to SQLite. 

## Calling the function

If we defined the query properly, any time we connect via SQLite, the engine will compile the C# code, notice the function and register it with SQLite.

![Calling C# table-valued function](http://www.querystorm.com/downloads/images/cs_tvf_call.png)

If we like, we can also use this function in joins. For example, if we have a table with transactions, and for each transaction has a column specifying a list of desired currency conversions, we can do it like so:

```sql
select 
	* 
from 
	transactions tr, 
	GetExchangeData(tr.date, tr.baseCurrency, tr.exchangeCurrencies)
```
Here we're joining each row in the transactions table with the results of calling the `GetExchangeData` function. For each row, we're passing in the values from the transactions row to the `GetExchangeData` function. 

It might seem strange to do a join this way, but it's just syntax sugar. Under the hood, SQLite sees table-valued functions as normal tables where parameters are just hidden columns. If you imagine that GetExchangeData is a normal table that has exchange rates, we can rewrite the query (and get the same result) like so:

```sql
select 
	* 
from 
	transactions tr 
	inner join GetExchangeData ged on
		ged.date = tr.date
		and ged.baseCurrency = tr.baseCurrency
		and get.symbols = tr.exchangeCurrencies
```
!!! Note
	The above query will actually run. QueryStorm's SQLite parser doesn't support this syntax so it shows false error squiggles, but the query will work just fine if you run it. 

## Caching
Caching is important when table joins are at play. We don't want to repeatedly hit a web service in the background if the function is called multiple times with the same parameters. That's why table-valued function results are cached, so that multiple calls **with the same arguments** only execute the table-valued function once and the result is cached and reused. 

The **caching scope is per-execution**, meaning that if a join results in multiple executions of the function with the same parameters, the function is only executed once, but if you run the query again it will indeed execute again.