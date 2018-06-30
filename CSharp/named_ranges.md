# Named ranges

Each named range in the workbook appears as a strongly typed global variable (technically it's a property, to be exact) that allows reading and writing the value in the named range. 

## Reading named ranges values
Depending on the value contained the named range (in Excel), the variable will be of the following type:

- number => `double`
- text => `string`
- date => `DateTime`
- time => `double` (fraction of the day, e.g. 6:00 AM => 0.25)
- boolean => `boolean`
- array (multiple cells selected) => `object []`

You can use the variable in your queries just like a normal variable, e.g.:

```csharp
return myNumericNamedRange + 50;
```

## Setting named range values
Setting the value of the variable will update the value in the named range in Excel. If for example, you have a named range that has a numeric value, you can update its value to 123 like so: 

```csharp
myNumericNamedRange = 123;
```

Note that you cannot set the variable to a value of a different type. If you need to use a different type, you must change the value manually in Excel or access the `Range` object explicitly:

```C#
Range(nameof(myNumericNamedRange)).Value = "a string value";
```  
