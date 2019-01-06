# Named ranges

Each named range in the workbook appears as a strongly typed global variable (technically it's a property) that allows reading and writing the value in the named range. 

## Reading named ranges values
Depending on the value of the named range, the variable that represents it will be of the corresponding type (formatting also plays a part for dates and currencies, though). You can use the variable in your queries just like a normal variable, e.g.:

```csharp
return myNumericNamedRange + 50;
```

## Setting named range values
Setting the value of the variable will update the value in Excel. If for example, you have a named range that has a numeric value, you can update its value like so: 

```csharp
myNumericNamedRange = 123;
```

Note that you cannot set the variable to a value of a different type. If you need to use a different type, you must change the value manually in Excel or access the `Range` object explicitly:

```C#
Range(nameof(myNumericNamedRange)).Value = "a string value";
```  
