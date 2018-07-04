# QueryStorm's VBA API

QueryStorm provides an API that you can consume from VBA. The API is rather simple and currently only allows running and modifying embedded queries. 

!!! Download
	You can download the sample workbook from [here](../demofiles/vba_api_demo.xlsm)!

## Working with embedded queries 
Embedded queries and scripts are accessed via the `GetQuery(path)` method. The returned query object has the following methods:
``` C#
//gets the contents of the embedded query
string GetContents();
//sets the contents of the embedded query
void SetContents(string value);

//Runs the query asynchronously
void RunAsync(string successCallback, string failCallback);
//Runs the query synchronously
Recordset Run();
```


### Running a query synchronously

When running the query synchronously, the results can be used in the next line of VBA code, but Excel will not be responsive to user input while the query is running.

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
    Set data = Query.Run()

    '...do something with data if necessary...
End Sub
```

### Running a query asynchronously

Queries can be run asyncronously using the `RunAsync` method. The method starts the background execution of the query and immediately returns. If you need to process the results, you can pass in the names of the two VBA subroutines that should be called upon successful/failed execution of the query.

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

The `successCallback` subroutine can have zero parameters if it doesn't care about the results. Alternatively, it can specify one parameter, and that parameter must be declared as `Variant` or `Recordset`, as it will be used to pass the results into the callback.

The `failCallback` subroutine must specify one parameter declared as `String`. This parameter will be used to pass the error message into the callback.

You can pass in `Nothing` or "" as a callback if you don't want to process success or failure.

Since embedded queries specify the engine configuration themselves (which you can see/change in the automation dialog), nothing about the connection needs to be specified. 
