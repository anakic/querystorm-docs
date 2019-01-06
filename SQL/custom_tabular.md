# Custom table-valued functions

In QueryStorm, you can write C# table-valued functions in order to expose external data to SQLite e.g. REST APIs, local files, active directory data, external databases... 

> You can download the demo workbook from [here](../demofiles/tvf_currencies.xlsx).

## Defining the function

In this example, I'll load data from an [exchange rates API](https://www.exchangeratesapi.io) and join it with data in an Excel table.

For a start, let's connect via C# and paste in the following function into the editor: 

```csharp
using Newtonsoft.Json.Linq;

///<summary>Gets the exchange rates for the specified currency on the specified date.</summary>
///<param name="date">The date for which to get the currency rates for</param>
[SQLFuncTabular]
public static IEnumerable<Rate> GetExchangeData(DateTime date, string baseCurrency = "EUR", string symbols = null)
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

We can, of course, call this function from C#, but we can also expose it to SQLite. To do so, we need to to do the following:

1. make sure it successfully compiles (has no compile-time errors)
1. decorate it with the `[SQLFuncTabular]` attribute 
1. embed it into the workbook

Once that's done, it will be available to SQLite. 

## Calling the function

When we connect via SQLite, it will compile the C# code in the workbook and register the function we defined.

![Calling C# table-valued function](../images/cs_tvf_call.png)

If we like, we can also use this function in joins. For example, let's assume we have a table with transactions, and each transaction has a *Date* column  and a *Target currencies* column. The `date` specifies when the transaction happened, and the `target currencies` specifies the currencies we need to convert into. 

![Currency](../images/tvf_currency.png)

We can join each transaction with the appropriate exchange rates like so:

```sql
select 
	* 
from 
	transactions tr, 
	GetExchangeData(tr.date, tr.baseCurrency, tr.exchangeCurrencies)
```

Here we're joining each row in the transactions table with the results of calling the `GetExchangeData` function. We're achieving the join by passing values from *transactions* rows into the `GetExchangeData` function. The resulting data looks something like this:

![Currency](../images/tvf_currency2.png)

It might seem strange to do a join this way, but it's just syntax sugar. Under the hood, SQLite sees table-valued functions as normal tables and their parameters as hidden columns of those tables. Imagine that `GetExchangeData` is a normal table that has all exchange rates ever (a better name for it might be simply *ExchangeRates*, though). With this in mind, we can rewrite the query like so:

```sql
select 
	* 
from 
	transactions tran
	inner join GetExchangeData exchage on
		exchage.date = tran.date
		and exchage.baseCurrency = tran.baseCurrency
		and exchage.symbols = tran.exchangeCurrencies
```
!!! Note
	The query works exactly the same as the previous one, only the syntax is different. 

## Caching
Caching is important when table joins are at play. We don't want to repeatedly hit a web service in the background if the function is called multiple times with the same parameters. That's why table-valued function results are cached so that multiple calls **with the same arguments** only execute the table-valued function once and the result is cached and reused. 

The **caching scope is per-execution**, meaning that if a join results in multiple executions of the function with the same parameters, the function is only executed once, but if you run the query again it will be re-executed.