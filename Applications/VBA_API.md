# QueryStorm's VBA API

QueryStorm provides an API that you can consume from VBA. The API is rather simple at the moment and will evolve as necessary. You can use it to run an embedded query, or a SQL command (against any engine). If you need to process the results, you can register success and error callback subroutines (by name). The callback approach is used because execution engines in QueryStorm are asynchronous.   

## A few examples
To get a sense of how the API is used, let's see some examples before moving on to the API reference.

> You can **download** a sample workbook with these examples **[here](https://www.querystorm.com/Downloads/vba_api_demo.xlsm)**!

### Example 1: running an Embedded query, and processing the results

```vba
Sub Button1_Click()
    'get the root API object
    Set api = Application.COMAddIns("QueryStorm").Object

    'get the embedded query we want to run
    Set q = api.GetQuery("Query1")

    'Run the embedded query, specifying the success and error callbacks. 
	'Pass in Nothing, or "" if no result processing is needed.
	'IMPORTANT NOTE: If the callbacks are not in a module (e.g. they're in a sheet) they 
	'need to be qualified, e.g. Sheet1.Yay, Sheet1.Boo
    Call q.Run("Yay", "Boo")
End Sub

Sub Yay(data As Variant)
    Set result = data.Fields.Item(0)
    MsgBox ("The result is " & result)
End Sub

Sub Boo(message As String)
    MsgBox ("Error: " & message)
End Sub
```

### Example 2: running SQL queries (non-embedded) with LiveMode engine
```vba 
Sub Button1_Click()

    'get the root API object
    Set api = Application.COMAddIns("QueryStorm").Object

    'create a runner object
    Set runner = api.CreateRunner()

    'register queries -> runners don't run queries directly, they give you promises
    Set promise1 = runner.AddQuery("select 123")
    Set promise2 = runner.AddQuery("select * from table1")
    
    'set callbacks using the promise objects
    promise1.OnSuccess("SuccessCallback1").OnError("ErrorCallback")
    promise2.OnSuccess("SuccessCallback2").OnError("ErrorCallback")
    
	'Open connection, run all queries (in this case 2 queries), close connection
	'Execution is done asynchronously in the background
    Call runner.Go
End Sub

Sub SuccessCallback1(data As Variant)
    'process as needed
End Sub

Sub SuccessCallback2(data As Variant)
    'process as needed
    'just here to illustrate that each promise can specify its own callback
End Sub

Sub ErrorCallback(message As String)
    MsgBox ("Error: " & message)
End Sub
```

### Example 3: running SQL query with SQL Server, also passing additional parameter to callback
``` vba
Sub Button1_Click()
    'get the root API object
    Set api = Application.COMAddIns("QueryStorm").Object
    
    'create an SQL Server runner
    Set runner = api.CreateRunnerMS("Server=.\sqlexpress; Database=AdventureWorks2012; Trusted_Connection=True;")
    
    'register table to copy to server as temp table
    runner.AddTable ("Table1")
    
    'register query
    Set promise = runner.AddQuery("select top {Sheet1!B1} * from production.product")

    'set query execution callbacks
    promise.OnSuccess("OnSuccess").OnError ("OnError")
   
    'also add another parameter that will get passed to callback
    promise.SetParameter ("Some dummy parameter")
    
    'run all queries (in this case, just 1 query)
    Call runner.Go
End Sub

'The success callback can specify 0, 1 or 2 parameters
'1st parameter is the data (ADO Recordset)
'2nd paramater is arbitraty data you can pass in using promise.SetParameter(...)
Sub OnSuccess(data As Variant, parameter As Variant)
    'process as needed
End Sub

'Error callback can specify 0, 1 or 2 parameters
'1st parameter is the error message (String)
'2nd paramater is arbitraty data you can pass in via the promise object
Sub OnError(message As String)
    MsgBox ("Error: " & message)
End Sub
```


----------

## API Reference

### The root API object
You can get a reference to the root QueryStorm API object as follows:
``` vba
Set api = Application.COMAddIns("QueryStorm").Object
```
This object exposes the methods shown below:
``` C#
//gets an embedded query by name
Query GetQuery(string name);

//creates a runner for running SQL commands in the Live mode
LiveRunner CreateRunner();

//creates a runner for running SQL commands with an Sql Server engine
SnapshotRunner CreateRunnerMS(string connectionString);

//creates a runner for running SQL commands with a PostreSQL engine
SnapshotRunner CreateRunnerPG(string connectionString);

//creates a runner for running SQL commands with a SQLite engine
SnapshotRunner CreateRunnerSL(string connectionString);

//creates a runner for running SQL commands with an Access engine
SnapshotRunner CreateRunnerAC(string connectionString);

//creates a runner for running SQL commands with a MySql engine
SnapshotRunner CreateRunnerMY(string connectionString);
```

### Running embedded queries
The `Query` object that is returned by `GetQuery(name)` **represents an embedded query**, and has the following methods:
``` C#
//gets the contents of the embedded query
string GetContents();

//sets the contents of the embedded query
void SetContents(string value);

//Runs the query
void Run(string successCallback, string failCallback);
```
The `Run` method is asynchronous, it starts the (background) execution of the query and immediately returns. If you need to process the results, pass in the names of the VBA subroutines that should be called upon successful/failed execution of the query.

The `successCallback` subroutine can have zero parameters if it doesn't care about the results. Alternatively, it can specify one parameter, and that parameter must be declared as `Variant` or `Recordset`, as it will be used to pass the results into the callback.

The `failCallback` subroutine must specify one parameter declared as `String`. This parameter will be used to pass the error message into the callback.

You can pass in `Nothing` or "" as a callback if you don't want to process success or failure.

Since embedded queries specify the engine configuration themselves (which you can see/change in the automation dialog), nothing about the connection needs to be specified. 

### Running non-embedded queries 
Usually, embedded queries are the way to go as far as automation is concerned. Embedded queries have their own connection data, but if we're executing arbitrary non-embedded queries, we need to specify and configure the connection using VBA. We do this by getting the appropriate *runner* from the root API object and configuring it.

There is a method for creating a runner for each type of engine. For example, here's how we would create a runner for a SqlServer database:
``` vba
'get the root API object
Set api = Application.COMAddIns("QueryStorm").Object

'create an SQL Server runner (notice the MS suffix of the method, MS = Ms Sql server)
Set runner = api.CreateRunnerMS("Server=.\sqlexpress; Database=AdventureWorks2012; Trusted_Connection=True;")
```
It's worth noting that its fine to have preprocessor expressions and directives in the queries that the runner will run. 

Runners execute a little bit differently than embedded queries. A runner can queue up several queries and, when ready, execute them all in one batch using a single connection. 

When adding a query to the runner, the runner will return a `promise` for the execution of that query. The promise object can be used to configure callbacks, and additional parameters.

Nothing gets executed until you call `Go` on the runner. Calling `Go` kicks off the background execution of the entire batch of queries. 

There are two types of runners that correspond to the two modes of querying in QueryStorm: LiveRunner and SnapshotRunner. They differ only slightly in their API.

#### 1. The LiveRunner type
LiveRunner is a runner specific to the LiveMode engine. It has the following methods:

``` C#
//scopeId: 0-ActiveSheet, 1-ActiveWorkbook, 2-AllOpenWorkbooks
void SetScope(int scopeId);

//schedule a query for execution, and return a promise
//that can be used to configure callbacks and other parameters
IResultPromise AddQuery(string query);

//Connect, run all scheduled queries (along with executing callbacks), disconnect
void Go();
``` 

#### 2. The SnapshotRunner type
`SnapshotRunner` object is used to run queries using an external engine. It has the following methods:
``` C#
//methods to set tables to copy to the destination server as temp tables 
void AddAllTables();
void AddTable(string table);
void AddTablesFromSheet(string sheetName);
void AddTablesFromWorkbook();
void RemoveTable(string table);

//schedule a query for execution, and return a promise
//that can be used to configure callbacks and other parameters
IResultPromise AddQuery(string query);

//Connect, run all scheduled queries (along with callbacks), disconnect
void Go();
``` 

> Note: it's called `SnapshotRunner` because it takes a snapshot of the data in Excel tables, and copies it into temp tables in an external database, it basically represents an external database engine.

#### 3. The ResultPromise type
When registering a query with `runner.AddQuery(queryText)`, a runner will return a *promise*. This *promise* can be used to configure callbacks (and more) for the query. The promise object has the following methods:
  
``` C#
ResultPromise OnSuccess(string successCallback);
ResultPromise OnError(string errorCallback);
ResultPromise SetParameter(object parameter);
ResultPromise StopOnError(bool shouldStopOnError);
``` 

> Note: The methods return the same promise object just to enable method chaining, e.g. `promise.OnSuccess('...').OnError(...)`

The success subroutine should specify 2 parameters. The first parameter is the data (ADO Recordset), the 2nd parameter is arbitrary data you can pass in via the promise object. It can optionally skip the second or both parameters.

The error subroutine can specify 2 parameters. The first parameter is the error message (String), the 2nd parameter is arbitrary data you can pass in via the promise object. It can optionally skip the second or both parameters.