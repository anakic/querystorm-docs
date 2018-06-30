# SQLite function reference
SQLite comes with many useful built in functions, but QueryStorm adds many useful functions of its own. Functions that start with an lowercase letter are SQLite's built in functions, while functions that start with an uppercase letter are QueryStorm's extensions. This document provides a reference of all the available functions. 

## Aggregate functions
### `avg`
Returns the average value of all non-NULL values within a group. String and BLOB values that do not look like numbers are interpreted as 0. The result of avg() is always a floating point value as long as at there is at least one non-NULL input even if all inputs are integers. The result of avg() is NULL if and only if there are no non-NULL inputs.

| Name | Type  | Description  |
|---|---|---|
| `value` | `numeric` | Value to aggregate. |

**Returns: `float`**



### `count`
Returns a count of the number of times that value is not NULL in a group. The count(*) function (with no arguments) returns the total number of rows in the group.

| Name | Type  | Description  |
|---|---|---|
| `X` | `numeric` | Value to aggregate. |

**Returns: `float`**



### `group_concat`
[Aggregate] Returns a string which is the concatenation of all non-NULL 'value' values. If parameter 'separator' is present then it is used as the separator between instances of 'value'. A comma (',') is used as the separator if Y is omitted. The order of the concatenated elements is arbitrary.

| Name | Type  | Description  |
|---|---|---|
| `value` | `any` | Value to aggregate. |
| `separator` | `string` | Separator. |

**Returns: `string`**



### `max`
[Aggregate] Returns the maximum value of all values in the group. The maximum value is the value that would be returned last in an ORDER BY on the same column. Aggregate max() returns NULL if and only if there are no non-NULL values in the group.

| Name | Type  | Description  |
|---|---|---|
| `value` | `numeric` | Value to aggregate. |

**Returns: `float`**



### `min`
[Aggregate] Returns the minimum non-NULL value of all values in the group. The minimum value is the first non-NULL value that would appear in an ORDER BY of the column. Aggregate min() returns NULL if and only if there are no non-NULL values in the group.

| Name | Type  | Description  |
|---|---|---|
| `value` | `numeric` | Value to aggregate. |

**Returns: `float`**



### `sum`
[Aggregate] Returns the sum of all non-NULL values in the group. If there are no non-NULL input rows then sum() returns NULL.

| Name | Type  | Description  |
|---|---|---|
| `value` | `numeric` | Value to aggregate. |

**Returns: `float`**



### `stdev`
[Aggregate] Returns the standard deviation of values in the group. Returns NULL if and only if there are no non-NULL values in the group.

| Name | Type  | Description  |
|---|---|---|
| `value` | `numeric` | Value to aggregate. |

**Returns: `float`**



### `variance`
[Aggregate] Returns the variance of values in the group. Returns NULL if and only if there are no non-NULL values in the group.

| Name | Type  | Description  |
|---|---|---|
| `value` | `numeric` | Value to aggregate. |

**Returns: `float`**



### `mode`
[Aggregate] Returns the value that appears most often in the group. Returns NULL if and only if there are no non-NULL values in the group.

| Name | Type  | Description  |
|---|---|---|
| `value` | `numeric` | Value to aggregate. |

**Returns: `float`**



### `median`
[Aggregate] Returns the median of values in the group. Returns NULL if and only if there are no non-NULL values in the group.

| Name | Type  | Description  |
|---|---|---|
| `value` | `numeric` | Value to aggregate. |

**Returns: `float`**



### `lower_quartile`
[Aggregate] Returns the lower quartile of values in the group. Returns NULL if and only if there are no non-NULL values in the group.

| Name | Type  | Description  |
|---|---|---|
| `value` | `numeric` | Value to aggregate. |

**Returns: `float`**



### `upper_quartile`
[Aggregate] Returns the upper quartile of values in the group. Returns NULL if and only if there are no non-NULL values in the group.

| Name | Type  | Description  |
|---|---|---|
| `value` | `numeric` | Value to aggregate. |

**Returns: `float`**



### `ElementAt`
Aggregate function. Returns the N-th value from [ValueSet] ordered by [OrderSet]. Returns NULL if N-th value not found.

| Name | Type  | Description  |
|---|---|---|
| `N` | `integer` | The 1-based index of the element to return. Positive values assumes ascending order, negative value means descending. |
| `ValueSet` | `any` | The set of values out of which to return the N-th element. |
| `OrderSet` | `any` | The set of values by which to order the elements of the set (without specifying the order, N-th element would mean nothing). |

**Returns: `any`**
#### Example
```sql
--Find the most populous city in each country (Negative N used to select MOST populous)
SELECT CountryCode, ElementAt(-1, Name, Population), ElementAt(-1, Population, Population) as Pop FROM city GROUP BY CountryCode ORDER BY Pop;

--Find the least populous city in each country (Positive N used to select LEAST populous)
SELECT CountryCode, ElementAt(1, Name, Population), ElementAt(1, Population, Population) as Pop FROM city GROUP BY CountryCode ORDER BY Pop;
```


### `total`
Provided as a convenient way to work around the design problems of sum() function. If there are no non-NULL input rows then sum() returns NULL but total() returns 0.0. Sum() will throw an 'integer overflow' exception if all inputs are integers or NULL and an integer overflow occurs at any point during the computation. Total() never throws an integer overflow.

| Name | Type  | Description  |
|---|---|---|
| `value` | `numeric` | Value to aggregate. |

**Returns: `float`**



## Math functions
### `abs`
Returns the absolute value of the numeric argument X. Abs(X) returns NULL if X is NULL. Abs(X) returns 0.0 if X is a string or blob that cannot be converted to a numeric value. If X is the integer -9223372036854775808 then abs(X) throws an integer overflow error since there is no equivalent positive 64-bit two complement value.

| Name | Type  | Description  |
|---|---|---|
| `X` | `numeric` | Signed value to convert to absolute value. |

**Returns: `float`**



### `cot`
Returns the cotangent of the numeric argument X.

| Name | Type  | Description  |
|---|---|---|
| `X` | `numeric` | Input value. |

**Returns: `float`**



### `coth`
Returns the hyperbolic cotangent of the numeric argument X.

| Name | Type  | Description  |
|---|---|---|
| `X` | `numeric` | Input value. |

**Returns: `float`**



### `cos`
Returns the cosine of the numeric argument X.

| Name | Type  | Description  |
|---|---|---|
| `X` | `numeric` | Input value. |

**Returns: `float`**



### `cosh`
Returns the hyperbolic cosine of the numeric argument X.

| Name | Type  | Description  |
|---|---|---|
| `X` | `numeric` | Input value. |

**Returns: `float`**



### `acos`
Returns the inverse cosine of the numeric argument X.

| Name | Type  | Description  |
|---|---|---|
| `X` | `numeric` | Input value. |

**Returns: `float`**



### `acosh`
Returns the inverse hyperbolic cosine of the numeric argument X.

| Name | Type  | Description  |
|---|---|---|
| `X` | `numeric` | Input value. |

**Returns: `float`**



### `sin`
Returns the sine of the numeric argument X.

| Name | Type  | Description  |
|---|---|---|
| `X` | `numeric` | Input value. |

**Returns: `float`**



### `sinh`
Returns the hyperbolic sine of the numeric argument X.

| Name | Type  | Description  |
|---|---|---|
| `X` | `numeric` | Input value. |

**Returns: `float`**



### `asin`
Returns the inverse sine of the numeric argument X.

| Name | Type  | Description  |
|---|---|---|
| `X` | `numeric` | Input value. |

**Returns: `float`**



### `asinh`
Returns the inverse hyperbolic sine of the numeric argument X.

| Name | Type  | Description  |
|---|---|---|
| `X` | `numeric` | Input value. |

**Returns: `float`**



### `tan`
Returns the tangent of the numeric argument X.

| Name | Type  | Description  |
|---|---|---|
| `X` | `numeric` | Input value. |

**Returns: `float`**



### `atan`
Returns the arctangent of the numeric argument X.

| Name | Type  | Description  |
|---|---|---|
| `X` | `numeric` | Input value. |

**Returns: `float`**



### `atanh`
Returns the inverse hyperbolic tangent of the numeric argument X.

| Name | Type  | Description  |
|---|---|---|
| `X` | `numeric` | Input value. |

**Returns: `float`**



### `atan2`
Returns the arc tangent in all four quadrants from specified values.

| Name | Type  | Description  |
|---|---|---|
| `Y` | `numeric` | Input value. |
| `X` | `numeric` | Input value. |

**Returns: `float`**



### `floor`
Returns the closest integer lower than or equal to X.

| Name | Type  | Description  |
|---|---|---|
| `X` | `float` | Input value. |

**Returns: `int`**



### `ceil`
Returns the closest integer greater than or equal to X.

| Name | Type  | Description  |
|---|---|---|
| `X` | `float` | Input value. |

**Returns: `int`**



### `degrees`
Converts the input value from radians to degrees.

| Name | Type  | Description  |
|---|---|---|
| `X` | `numeric` | Input value. |

**Returns: `float`**



### `radians`
Converts the input value from degrees to radians.

| Name | Type  | Description  |
|---|---|---|
| `degrees` | `numeric` | Input value. |

**Returns: `float`**



### `log`
Returns the natural logarithm (base e) of the numeric input value.

| Name | Type  | Description  |
|---|---|---|
| `X` | `numeric` | Input value. |

**Returns: `float`**



### `log10`
Returns the logarithm (base 10) of the numeric input value.

| Name | Type  | Description  |
|---|---|---|
| `X` | `numeric` | Input value. |

**Returns: `float`**



### `sign`
Returns -1 if the input number is negative, 1 if the input number is positive and 0 if the input number is 0.

| Name | Type  | Description  |
|---|---|---|
| `X` | `numeric` | Input value. |

**Returns: `int`**



### `power`
Returns a specified number raised to the specified power.

| Name | Type  | Description  |
|---|---|---|
| `X` | `numeric` | Number. |
| `Y` | `numeric` | Power. |

**Returns: `float`**



### `sqrt`
Returns the square root of the specified number.

| Name | Type  | Description  |
|---|---|---|
| `X` | `numeric` | Number. |

**Returns: `float`**



### `square`
Returns the square of the specified number.

| Name | Type  | Description  |
|---|---|---|
| `X` | `numeric` | Number. |

**Returns: `float`**



### `max`
Returns the argument with the maximum value, or return NULL if any argument is NULL. The multi-argument max() function searches its arguments from left to right for an argument that defines a collating function and uses that collating function for all string comparisons. If none of the arguments to max() define a collating function, then the BINARY collating function is used. Note that max() is a simple function when it has 2 or more arguments but operates as an aggregate function if given only a single argument.

| Name | Type  | Description  |
|---|---|---|
| `X1...XN` | `numeric` | Input values. |

**Returns: `float`**



### `min`
Returns the argument with the minimum value. The multi-argument min() function searches its arguments from left to right for an argument that defines a collating function and uses that collating function for all string comparisons. If none of the arguments to min() define a collating function, then the BINARY collating function is used. Note that min() is a simple function when it has 2 or more arguments but operates as an aggregate function if given only a single argument.

| Name | Type  | Description  |
|---|---|---|
| `X1...XN` | `numeric` | Input values. |

**Returns: `float`**



### `round`
Returns a floating-point value X rounded to Y digits to the right of the decimal point. If the Y argument is omitted, it is assumed to be 0.

| Name | Type  | Description  |
|---|---|---|
| `X` | `float` | Input number. |
| `Y` | `int` | Desired decimal places. |

**Returns: `float`**



### `random`
Returns an integer value between Min and Max. Both limits are inclusive.

| Name | Type  | Description  |
|---|---|---|
| `Min` | `int` | The minimum possible value to be generated. |
| `Max` | `int` | The maximum possible value to be generated. |

**Returns: `int`**



### `RandomGauss`
The RandomGauss function returns a random number generated according to a normal (gaussian) distribution.

| Name | Type  | Description  |
|---|---|---|
| `mean` | `numeric` | The mean value of the distribution. |
| `sigma` | `numeric` | The standard deviation of the distribution. |

**Returns: `float`**



### `RandomGauss`
The RandomGauss function returns a random number generated according to a normal (gaussian) distribution.


**Returns: `float`**



### `random`
Returns a pseudo-random integer between -9223372036854775808 and +9223372036854775807.


**Returns: `int`**



### `pi`
Returns pi.


**Returns: `float`**



### `randomblob`
Return an N-byte blob containing pseudo-random bytes. If N is less than 1 then a 1-byte random blob is returned.

| Name | Type  | Description  |
|---|---|---|
| `N` | `int` | Desired length of the output blob. |

**Returns: `blob`**



## Core functions
### `char`
Returns a string composed of characters having the unicode code point values of integers X1 through XN, respectively.

| Name | Type  | Description  |
|---|---|---|
| `X1..XN` | `int` | Chars to concatenate. |

**Returns: `string`**



### `coalesce`
Returns a copy of its first non-NULL argument, or NULL if all arguments are NULL. Coalesce() must have at least 2 arguments.

| Name | Type  | Description  |
|---|---|---|
| `X1...XN` | `any` | Values. |

**Returns: `any`**



### `glob`
Equivalent to the expression 'Y GLOB X'. Note that the X and Y arguments are reversed in the glob() function relative to the infix GLOB operator.

| Name | Type  | Description  |
|---|---|---|
| `X` | `string` | Input string. |
| `Y` | `string` | Pattern. |

**Returns: `bool`**



### `ifnull`
Returns a copy of its first non-NULL argument, or NULL if both arguments are NULL. Ifnull() must have exactly 2 arguments. The ifnull() function is equivalent to coalesce() with two arguments.

| Name | Type  | Description  |
|---|---|---|
| `X` | `any` | Value1. |
| `Y` | `any` | Value2. |

**Returns: `any`**



### `instr`
Finds the first occurrence of string Y within string X and returns the number of prior characters plus 1, or 0 if Y is nowhere found within X. Or, if X and Y are both BLOBs, then instr(X,Y) returns one more than the number bytes prior to the first occurrence of Y, or 0 if Y does not occur anywhere within X. If both arguments X and Y to instr(X,Y) are non-NULL and are not BLOBs then both are interpreted as strings. If either X or Y are NULL in instr(X,Y) then the result is NULL.

| Name | Type  | Description  |
|---|---|---|
| `X` | `string` | Input text. |
| `Y` | `string` | Substring to search for. |

**Returns: `int`**



### `hex`
Interprets its argument as a BLOB and returns a string which is the upper-case hexadecimal rendering of the content of that blob.

| Name | Type  | Description  |
|---|---|---|
| `X` | `any` | Input. |

**Returns: `string`**



### `last_insert_rowid`
Returns the ROWID of the last row insert from the database connection which invoked the function.


**Returns: `int`**



### `length`
For a string value X, returns the number of characters (not bytes) in X prior to the first NUL character. Since SQLite strings do not normally contain NUL characters, the length(X) function will usually return the total number of characters in the string X. For a blob value X, length(X) returns the number of bytes in the blob. If X is NULL then length(X) is NULL. If X is numeric then length(X) returns the length of a string representation of X.

| Name | Type  | Description  |
|---|---|---|
| `X` | `string` | Input string. |

**Returns: `int`**



### `lower`
Returns a copy of string X with all ASCII characters converted to lower case. The default built-in lower() function works for ASCII characters only. To do case conversions on non-ASCII characters, load the ICU extension.

| Name | Type  | Description  |
|---|---|---|
| `X` | `string` | Input string. |

**Returns: `string`**



### `ltrim`
Returns a string formed by removing any and all characters that appear in Y from the left side of X. If the Y argument is omitted, ltrim(X) removes spaces from the left side of X.

| Name | Type  | Description  |
|---|---|---|
| `X` | `string` | Input string. |
| `Y` | `string` | String with characters to remove from input string. |

**Returns: `string`**



### `nullif`
Returns its first argument if the arguments are different and NULL if the arguments are the same. The nullif(X,Y) function searches its arguments from left to right for an argument that defines a collating function and uses that collating function for all string comparisons. If neither argument to nullif() defines a collating function then the BINARY is used.

| Name | Type  | Description  |
|---|---|---|
| `X` | `any` | Argument1. |
| `Y` | `any` | Argument2. |

**Returns: `any`**



### `printf`
The printf(FORMAT,...) SQL function works like the printf() function from the standard C library. The first argument is a format string that specifies how to construct the output string using values taken from subsequent arguments. If the FORMAT argument is missing or NULL then the result is NULL. The %n format is silently ignored and does not consume an argument. The %p format is an alias for %X. The %z format is interchangeable with %s. If there are too few arguments in the argument list, missing arguments are assumed to have a NULL value, which is translated into 0 or 0.0 for numeric formats or an empty string for %s.

| Name | Type  | Description  |
|---|---|---|
| `FORMAT` | `string` | Argument1. |
| `X1...XN` | `any` | Argument2. |

**Returns: `string`**



### `quote`
Returns the text of an SQL literal which is the value of its argument suitable for inclusion into an SQL statement. Strings are surrounded by single-quotes with escapes on interior quotes as needed. BLOBs are encoded as hexadecimal literals. Strings with embedded NUL characters cannot be represented as string literals in SQL and hence the returned string literal is truncated prior to the first NUL.

| Name | Type  | Description  |
|---|---|---|
| `X` | `any` | Value. |

**Returns: `string`**



### `replace`
Returns a string formed by substituting string Z for every occurrence of string Y in string X. The BINARY collating sequence is used for comparisons. If Y is an empty string then return X unchanged. If Z is not initially a string, it is cast to a UTF-8 string prior to processing.

| Name | Type  | Description  |
|---|---|---|
| `X` | `string` | Input string. |
| `Y` | `string` | Search string. |
| `Z` | `string` | Replacement string. |

**Returns: `string`**



### `rtrim`
Returns a string formed by removing any and all characters that appear in Y from the right side of X. If the Y argument is omitted, rtrim(X) removes spaces from the right side of X.

| Name | Type  | Description  |
|---|---|---|
| `X` | `string` | Input string. |
| `Y` | `string` | String with characters to remove from input string. |

**Returns: `string`**



### `sqlite_version`
Returns the version string for the SQLite library that is running.


**Returns: `string`**



### `substr`
Returns a substring of input string X that begins with the Y-th character and which is Z characters long. If Z is omitted then substr(X,Y) returns all characters through the end of the string X beginning with the Y-th. The left-most character of X is number 1. If Y is negative then the first character of the substring is found by counting from the right rather than the left. If Z is negative then the abs(Z) characters preceding the Y-th character are returned. If X is a string then characters indices refer to actual UTF-8 characters. If X is a BLOB then the indices refer to bytes.

| Name | Type  | Description  |
|---|---|---|
| `X` | `string` | Input string. |
| `Y` | `int` | Start location (1-based). |
| `Z` | `int` | Length of substring. |

**Returns: `string`**



### `trim`
Returns a string formed by removing any and all characters that appear in Y from both ends of X. If the Y argument is omitted, trim(X) removes spaces from both ends of X.

| Name | Type  | Description  |
|---|---|---|
| `X` | `string` | Input string. |
| `Y` | `string` | String with characters to remove from input string. |

**Returns: `string`**



### `typeof`
Returns a string that indicates the datatype of the expression X: 'null', 'integer', 'real', 'text', or 'blob'.

| Name | Type  | Description  |
|---|---|---|
| `X` | `any` | Value. |

**Returns: `string`**



### `unicode`
Returns the numeric unicode code point corresponding to the first character of the string X. If the argument to unicode(X) is not a string then the result is undefined.

| Name | Type  | Description  |
|---|---|---|
| `X` | `string` | Value. |

**Returns: `string`**



### `upper`
Returns a copy of input string X in which all lower-case ASCII characters are converted to their upper-case equivalent.

| Name | Type  | Description  |
|---|---|---|
| `X` | `string` | Input string. |

**Returns: `string`**



## DateTime functions
### `date`
Returns the date in this format: YYYY-MM-DD.

| Name | Type  | Description  |
|---|---|---|
| `input` | `string` | The string representation of the date (IS0-8601) or 'now'. |
| `modifier1...modifierN` | `string` | Modifiers are applied to the input from left to right, order is important. Examples of valid modifiers: 'start of month','+1 month','-1 day'. |

**Returns: `string`**
#### Example
```sql
--Compute the current date.
SELECT date('now');

--Compute the last day of the current month.
SELECT date('now','start of month','+1 month','-1 day');
```


### `time`
Returns the time as HH:MM:SS.

| Name | Type  | Description  |
|---|---|---|
| `input` | `string` | The string representation of the date (IS0-8601) or 'now'. |
| `modifier1...modifierN` | `string` | Modifiers are applied to the input from left to right, order is important. Examples of valid modifiers: 'start of month','+1 month','-1 day'. |

**Returns: `string`**



### `datetime`
Returns the datetime in the format YYYY-MM-DD HH:MM:SS.

| Name | Type  | Description  |
|---|---|---|
| `input` | `string` | The string representation of the date (IS0-8601) or 'now'. |
| `modifier1...modifierN` | `string` | Modifiers are applied to the input from left to right, order is important. Examples of valid modifiers: 'start of month','+1 month','-1 day'. |

**Returns: `string`**
#### Example
```sql
--Compute the date of the first Tuesday in October for the current year.
SELECT date('now','start of year','+9 months','weekday 2');
```


### `julianday`
Returns the Julian day - the number of days since noon in Greenwich on November 24, 4714 B.C. (Proleptic Gregorian calendar).

| Name | Type  | Description  |
|---|---|---|
| `input` | `string` | The string representation of the date (IS0-8601) . |
| `modifier1...modifierN` | `string` | Modifiers are applied to the input from left to right, order is important. Examples of valid modifiers: 'start of month','+1 month','-1 day'. |

**Returns: `string`**
#### Example
```sql
--Compute the number of days since the signing of the US Declaration of Independence.
SELECT julianday('now') - julianday('1776-07-04');
```


### `strftime`
The strftime() routine returns the date formatted according to the format string specified as the first argument.

| Name | Type  | Description  |
|---|---|---|
| `format` | `string` | Valid formats: %d-day of month, %f-fractional seconds: SS.SSS, %H-hour, %j-day of year, %J-Julian day number, %m-month, %M-minute, %s-seconds since 1970-01-01, %S-seconds, %w-day of week 0-6 (Sunday=0), %W-week of year, %Y-year. |
| `timestring` | `string` | The string representation of the date (IS0-8601) or 'now'. |
| `modifier1...modifierN` | `string` | Modifiers are applied to the input from left to right, order is important. Examples of valid modifiers: 'start of month','+1 month','-1 day'. |

**Returns: `string`**
#### Example
```sql
--Compute the current unix timestamp.
SELECT strftime('%s','now');

--Compute the number of seconds since a particular moment in 2004:
SELECT strftime('%s','now') - strftime('%s','2004-01-01 02:34:56');
```


### `Now`
A string that represents the current date/time. (Wraps .NET's DateTime.Now property).


**Returns: `string`**



### `Today`
A string that represents the current date. (Wraps .NET's DateTime.Today property).


**Returns: `string`**



### `Utc`
Converts the specified date to Coordinated Universal Time (UTC).

| Name | Type  | Description  |
|---|---|---|
| `input` | `string` | Input datetime. |

**Returns: `string`**



### `Year`
Gets the year component of the specified datetime.

| Name | Type  | Description  |
|---|---|---|
| `input` | `string` | Input datetime. |

**Returns: `string`**



### `Month`
Gets the month component of the specified datetime.

| Name | Type  | Description  |
|---|---|---|
| `input` | `string` | Input datetime. |

**Returns: `string`**



### `Day`
Gets the day component of the specified datetime.

| Name | Type  | Description  |
|---|---|---|
| `input` | `string` | Input datetime. |

**Returns: `string`**



### `DayOfYear`
Gets the day of year of the specified datetime.

| Name | Type  | Description  |
|---|---|---|
| `input` | `string` | Input datetime. |

**Returns: `string`**



### `DayOfWeek`
Gets the day of week of the specified datetime.

| Name | Type  | Description  |
|---|---|---|
| `input` | `string` | Input datetime. |

**Returns: `string`**



### `Hour`
Gets the hour component of the specified datetime.

| Name | Type  | Description  |
|---|---|---|
| `input` | `string` | Input datetime. |

**Returns: `string`**



### `Minute`
Gets the minute component of the specified datetime.

| Name | Type  | Description  |
|---|---|---|
| `input` | `string` | Input datetime. |

**Returns: `string`**



### `Second`
Gets the second component of the specified datetime.

| Name | Type  | Description  |
|---|---|---|
| `input` | `string` | Input datetime. |

**Returns: `string`**



### `Millisecond`
Gets the millisecond component of the specified datetime.

| Name | Type  | Description  |
|---|---|---|
| `input` | `string` | Input datetime. |

**Returns: `string`**



### `FormatDateTime`
Converts a datetime value into a string according to desired format. Remark: The 'Format' function cannot be used with dates because SQLite internally simulates datetime with string (ISO) or float (JulianDay).

| Name | Type  | Description  |
|---|---|---|
| `input` | `string` | Input datetime. |
| `format` | `string` | The desired output format. |

**Returns: `string`**



### `OADate`
Converts a datetime string into the format Excel uses internally which is a double precision floating point number similar to Julian date but with a different starting point. (OADate is short for OLE Automation date). This function is used for overwriting datetime columns.

| Name | Type  | Description  |
|---|---|---|
| `input` | `string` | Input datetime. Can be a string(ISO) or double(julianday). |
| `format` | `double` | The OADate representation of the date. |

**Returns: `string`**



### `ParseDateTime`
Converts a datetime string with a custom format to an ISO formatted date string (SQLite simulates datetime with JulianDate and/or ISO8601 string).

| Name | Type  | Description  |
|---|---|---|
| `input` | `string` | The datetime string. |
| `format` | `string` | The format of the datetime string. |

**Returns: `string`**



## Text functions
### `proper`
Converts a string to proper case.

| Name | Type  | Description  |
|---|---|---|
| `text` | `string` | The input text. |

**Returns: `string`**



### `IsMatch`
Tests whether a regular expression pattern is matched inside a given input string. (Wrapper around .NET’s Regex.IsMatch function).

| Name | Type  | Description  |
|---|---|---|
| `input` | `string` | The input string. |
| `pattern` | `string` | The pattern to search for. |

**Returns: `string`**
#### Example
```sql
-- find rows where the FirstName contains a double letter, e.g. Terrence or Jessica
Select * from Contacts where ismatch(FirstName, '(\w)\1')
```


### `Concat`
Concatenates strings with the provided separator. Important distinction to regular concatenation (||) - NULL is treated as an empty string.

| Name | Type  | Description  |
|---|---|---|
| `Separator` | `string` | The input string. |
| `arg1..N` | `string` | The values to concatenate. |

**Returns: `string`**
#### Example
```sql
Select *, Concat(', ', FirstName, MiddleInitial, LastName) from Contacts
```


### `RegexReplace`
Replaces substrings that match a regular expression pattern with a replacement string. (Wrapper around .NET’s Regex.Replace method except that it can take multiple pairs of pattern/replacement arguments).

| Name | Type  | Description  |
|---|---|---|
| `input` | `string` | The input string. |
| `patternN` | `string` | The pattern to search for. Notes: Can reference its own groups (e.g. (\w)\1 will match double letters: \w matches a letter, the parenthesees capture it in a group and \1 references the first group). Multiple pairs of pattern/replacement can be specified. |
| `replacementN` | `string` | The replacement pattern. Notes: You can reference captured groups by index (e.g. $1) or by name (e.g. ${groupName}). multiple pairs of pattern/replacement can be specified. |

**Returns: `string`**
#### Example
```sql
-- RegexReplace is a wrapper arrount .NET's Regex.Replace method: 
-- https://msdn.microsoft.com/en-us/library/ewy2t5e0%28v=vs.110%29.aspx

-- Example 1: Separate first name and last by casing

--a) using implicit groups
select RegexReplace('JohnSmith', '([a-z])([A-Z])', '$1 $2');

--b) using named groups
select RegexReplace('JohnSmith', '(?<g1>[a-z])(?<g2>[A-Z])', '${g1} ${g2}');


-- Example 2: returns ab1122c33d
select RegexReplace(
	'aabb12cc3d',		-- some input string
	'\d', '$0$0', 		-- duplicate each digit (e.g. 12 -> 1122)  
	'([A-z])\1', '$1')	-- replace double letters with a single letter (e.g. aabb -> ab) 
	as res;
```


### `RegexExtract`
Returns the matched part of the input string.

| Name | Type  | Description  |
|---|---|---|
| `input` | `string` | The input string. |
| `pattern` | `string` | The pattern to search for. Notes: Can reference its own groups (e.g. (\w)\1 will match double letters: \w matches a letter, the parenthesees capture it in a group and \1 references the first group). Multiple pairs of pattern/replacement can be specified. |

**Returns: `string`**



### `RegexExtract`
Creates a new string with values captured from the input string.

| Name | Type  | Description  |
|---|---|---|
| `input` | `string` | The input string. |
| `pattern` | `string` | The pattern to search for. Notes: Can reference its own groups (e.g. (\w)\1 will match double letters: \w matches a letter, the parenthesees capture it in a group and \1 references the first group). Multiple pairs of pattern/replacement can be specified. |
| `outputTemplate` | `string` | The output template that will be filled with values from matched groups. Notes: You can reference captured groups by index (e.g. $1) or by name (e.g. ${groupName}). |

**Returns: `string`**



### `Format`
Populates a string template with (formatted) parameters. (Wrapper around .NET's String.Format method).

| Name | Type  | Description  |
|---|---|---|
| `template` | `string` | The template to which to insert the values. |
| `value1...valueN` | `any` | Values to format and insert into the placeholders |

**Returns: `string`**
#### Example
```sql
--Example1: construct a text message with formatted parameters
-- Sample output: Hello John McDoughal, your next haircut will be in the month of April.
select
    format('Hello {0} {1}, your next haircut will be in the month of {2:MMMM}.', FirstName, LastName, HaircutDate) as greetingMessage
from
    Contacts;
        
-- Example2: left-pad a name to a minimum of 20 characters
-- Output: '                John'
select format('{0,20}', 'John') from Contacts;
        
-- Example3: right-pad a name to a minimum of 20 characters
-- Output: 'John                '
select format('{0,-20}', 'John') from Contacts;
        
```


### `replicate`
Replicates the input text the specified number of times.

| Name | Type  | Description  |
|---|---|---|
| `text` | `string` | The input text. |
| `n` | `number` | Time to repeat the input text. |

**Returns: `string`**



### `leftstr`
Returns a substring of the input text consisting of the n leftmost characters.

| Name | Type  | Description  |
|---|---|---|
| `text` | `string` | The input text. |
| `n` | `number` | Substring length. |

**Returns: `string`**



### `rightstr`
Returns a substring of the input text consisting of the n rightmost characters.

| Name | Type  | Description  |
|---|---|---|
| `text` | `string` | The input text. |
| `n` | `number` | Substring length. |

**Returns: `string`**



### `padl`
Returns the input string padded from the left by spaces, to the minimum total length of n.

| Name | Type  | Description  |
|---|---|---|
| `text` | `string` | The input text. |
| `n` | `number` | The minimal total lenght. |

**Returns: `string`**



### `padr`
Returns the input string padded from the right by spaces, to the minimum total length of n.

| Name | Type  | Description  |
|---|---|---|
| `text` | `string` | The input text. |
| `n` | `number` | The minimal total lenght. |

**Returns: `string`**



### `padc`
Returns the input string padded from the right and left by spaces, to the minimum total length of n.

| Name | Type  | Description  |
|---|---|---|
| `text` | `string` | The input text. |
| `n` | `number` | The minimal total lenght. |

**Returns: `string`**



### `reverse`
Returns the reversed input string.

| Name | Type  | Description  |
|---|---|---|
| `text` | `string` | The input text. |

**Returns: `string`**



### `charindex`
Returns the 1-based index at which the specified searchString appears in the input text.

| Name | Type  | Description  |
|---|---|---|
| `searchText` | `string` | The string to search for. |
| `text` | `string` | The input text. |

**Returns: `int`**



### `strfilter`
Returns all the characters of the input string consisting that appear in the filter.

| Name | Type  | Description  |
|---|---|---|
| `text` | `string` | The input text. |
| `filter` | `string` | The filter characters. |

**Returns: `int`**



### `difference`
Computes the number of different characters between the soundex value of 2 strings

| Name | Type  | Description  |
|---|---|---|
| `textA` | `string` | First text in comparison. |
| `textB` | `string` | Second text in comparison. |

**Returns: `string`**



### `LevDist`
Computes the difference between two texts i.e. the number of character edits required to change one text into the other. Used for fuzzy text search. NOTE: The algorithm is Damerau-Levenshtein.

| Name | Type  | Description  |
|---|---|---|
| `textA` | `string` | First text in comparison. |
| `textB` | `string` | Second text in comparison. |
| `maxdiff` | `int` | The maximum allowed diff - specifying this can improve performance. For diffs greater than maxdiff, the function will return return int.MaxValue (2147483647) |

**Returns: `int`**
#### Example
```sql
/*Find possible typos in a table (pairs of people with different but very similar names). 
Computationally intensive; will take a while to complete for large tables (e.g. 10k rows or more).*/
SELECT
	a.FirstName || a.LastName as [Person A], b.FirstName || b.LastName as [Person B],
	LevDist(a.FirstName || a.LastName, b.FirstName || b.LastName, 2) diff,
	a.__address, b.__address
FROM
	Person a inner join Person b on a.__row < b.__row and diff BETWEEN 1 and 2
```


### `LevDist`
Measures the difference between two texts - the number of character edits (insertion/deletion/replacement/transposition) required to change one text into the other. Used for fuzzy text search. NOTE: To be precise, the algorithm used is not regular Levenshtein, but Damerau-Levenshtein (treats transposition as a single change)! 

| Name | Type  | Description  |
|---|---|---|
| `textA` | `string` | First text in comparison. |
| `textB` | `string` | Second text in comparison. |

**Returns: `int`**
#### Example
```sql
/*Find possible typos in a table (pairs of people with different but very similar names). 
Computationally intensive; will take a while to complete for large tables (e.g. 10k rows or more).*/
SELECT
	a.FirstName || a.LastName as [Person A], b.FirstName || b.LastName as [Person B],
	LevDist(a.FirstName || a.LastName, b.FirstName || b.LastName, 2) diff,
	a.__address, b.__address
FROM
	Person a inner join Person b on a.__row < b.__row and diff BETWEEN 1 and 2
```


### `EqNoCase`
Case insensitive comparison of two parameters; returns 1 if their string representations are the same (ignoring case) and 0 otherwise.

| Name | Type  | Description  |
|---|---|---|
| `strA` | `any` | Input string A. |
| `strB` | `any` | Input string B. |

**Returns: `bool`**



## General functions
### `NewId`
Returns a new Guid. (Wraps .NET's Guid.New).


**Returns: `string`**



### `XPath`
Evaluates XPath expressions on specified xml text.

| Name | Type  | Description  |
|---|---|---|
| `xml` | `string` | The input xml string. |
| `xpath` | `any` | The xpath expression to evaulate. |

**Returns: `any`**



### `XPath`
Evaluates XPath expressions on specified xml text.

| Name | Type  | Description  |
|---|---|---|
| `xml` | `string` | The input xml string. |
| `xpath` | `any` | The xpath expression to evaulate. |
| `separator` | `any` | The string that will be used to separate values (if multiple values are returned). |

**Returns: `any`**



### `Eval`
Evaluates a C#-like expression. Can be used to call .NET functionality and manipulate Excel objects (e.g. modify formatting of table rows).

| Name | Type  | Description  |
|---|---|---|
| `expression` | `string` | The expression to evaluate. Can reference parameters by index (0-based). |
| `parameter1..N` | `any` | Zero or more parameters for the expression. |

**Returns: `any`**
#### Example
```sql
--Add a random number of months to current date
select eval('DateTime.Parse($0).AddMonths($1)', date(), random(1,200))
```


### `GpsDist`
Distance in meters between two points on the Earth's surface that are specified with GPS coordinates.

| Name | Type  | Description  |
|---|---|---|
| `Lat1` | `float` | Latitude of point1. |
| `Lon1` | `float` | Longitude of point1. |
| `Lat2` | `float` | Latitude of point2. |
| `Lon2` | `float` | Longitude of point2. |

**Returns: `float`**



## Excel functions
### `Group_address`
Gets a custom string identifying a group of rows. This string can be used to apply formatting in bulk to all the rows in a group (e.g. using the SetBackgroundColor function). This makes the operation much faster since formatting is applied to multiple rows at once, instead of one by one.

| Name | Type  | Description  |
|---|---|---|
| `Address` | `string` | Address of the row to include in group. |

**Returns: `any`**



### `GetCellValue`
Gets the value of the cell with the specified address.

| Name | Type  | Description  |
|---|---|---|
| `Address` | `string` | Address of the cell whose values is to be returned. Most likely will be the hidden __address column that each table in QueryStorm has. Can be named range, qualified ('e.g. Sheet1!A4') or unqualified (e.g. 'A4') cell address. |

**Returns: `any`**



### `GetCellFormula`
Gets the formula in the cell with the specified address.

| Name | Type  | Description  |
|---|---|---|
| `Address` | `string` | Address of the cell whose formula is to be returned. Most likely will be the hidden __address column that each table in QueryStorm has. Can be named range, qualified ('e.g. Sheet1!A4') or unqualified (e.g. 'A4') cell address. |

**Returns: `any`**



### `SelectRange`
Selects the ranges with the specified addresses and returns a string with all selected addresses (IMPORTANT NOTE: this is an aggregate function for performance reasons)

| Name | Type  | Description  |
|---|---|---|
| `Address` | `string` | Address of the cell that is to be selected. Most likely will be the hidden column field that each table in QueryStorm has. Can be named range, qualified ('e.g. Sheet1!A4') or unqualified (e.g. 'A4') cell address. |

**Returns: `string`**
#### Example
```sql
-- select all people (their rows in Excel) in the persons table who are older than 30 
select SelectRange(__address) from Person where DateOfBirth < date('now', '-30 years')
```


### `SetCellValue`
Sets the value of the cell with the specified address to the specified value.

| Name | Type  | Description  |
|---|---|---|
| `Address` | `string` | Address of the cell whose values is to be set. Most likely will be the hidden __address column that each table in QueryStorm has. Can be named range, qualified ('e.g. Sheet1!A4') or unqualified (e.g. 'A4') cell address. |
| `Value` | `string` | Value to write into the specified cell. |

**Returns: `any`**



### `SetCellFormula`
Sets the formula of the cell with the specified address.

| Name | Type  | Description  |
|---|---|---|
| `Address` | `string` | Address of the cell whose formula is to be set. Most likely will be the hidden __address column that each table in QueryStorm has. Can be named range, qualified ('e.g. Sheet1!A4') or unqualified (e.g. 'A4') cell address. |
| `Formula` | `string` | The formula to write into the specified cell. |

**Returns: `any`**



### `SetBackgroundColor`
Sets the interior color of the cell with the specified address to the specified value.

| Name | Type  | Description  |
|---|---|---|
| `Address` | `string` | Address of the cell whose color is to be set. Most likely will be the hidden __address column that each table in QueryStorm has. Can be named range, qualified ('e.g. Sheet1!A4') or unqualified (e.g. 'A4') cell address. |
| `ColorName` | `string` | The name of the color to set. |

**Returns: `any`**
#### Example
```sql
SELECT
	CASE JobTitle
		WHEN 'Chief Executive Officer' THEN SetBackgroundColor(__address, 'red')
		WHEN 'Vice President of Engineering' THEN SetBackgroundColor(__address, 'yellow')
		WHEN 'Research and Development Manager' THEN SetBackgroundColor(__address, 'teal')
		ELSE ClearBackgroundColor(__address)
	END
FROM
	Employee
```


### `SetFontColor`
Sets the interior color of the cell with the specified address to the specified value.

| Name | Type  | Description  |
|---|---|---|
| `Address` | `string` | Address of the cell whose color is to be set. Most likely will be the hidden __address column that each table in QueryStorm has. Can be named range, qualified ('e.g. Sheet1!A4') or unqualified (e.g. 'A4') cell address. |
| `ColorName` | `string` | The name of the color to set. |

**Returns: `any`**
#### Example
```sql
SELECT
	CASE JobTitle
		WHEN 'Chief Executive Officer' THEN SetFontColor(__address, 'red')
		WHEN 'Vice President of Engineering' THEN SetFontColor(__address, 'yellow')
		WHEN 'Research and Development Manager' THEN SetFontColor(__address, 'teal')
		ELSE SetFontColor(__address, 'Black')
	END
FROM
	Employee
```


### `SetBackgroundColor`
Sets the interior color of the cell with the specified address to the specified value.

| Name | Type  | Description  |
|---|---|---|
| `Address` | `string` | Address of the cell whose color is to be set. Most likely will be the hidden __address column that each table in QueryStorm has. Can be named range, qualified ('e.g. Sheet1!A4') or unqualified (e.g. 'A4') cell address. |
| `Red` | `int` | The red component of the color to set (0-255). |
| `Green` | `int` | The green component of the color to set (0-255). |
| `Blue` | `int` | The blue component of the color to set (0-255). |

**Returns: `any`**
#### Example
```sql
SELECT
	CASE JobTitle
		WHEN 'Chief Executive Officer' THEN SetBackgroundColor(__address, 255, 0, 0)
		WHEN 'Vice President of Engineering' THEN SetBackgroundColor(__address, 255, 255, 0)
		WHEN 'Research and Development Manager' THEN SetBackgroundColor(__address, 0, 255, 255)
		ELSE ClearBackgroundColor(__address)
	END
FROM
	Employee
```


### `SetFontColor`
Sets the interior color of the cell with the specified address to the specified value.

| Name | Type  | Description  |
|---|---|---|
| `Address` | `string` | Address of the cell whose color is to be set. Most likely will be the hidden __address column that each table in QueryStorm has. Can be named range, qualified ('e.g. Sheet1!A4') or unqualified (e.g. 'A4') cell address. |
| `Red` | `int` | The red component of the color to set (0-255). |
| `Green` | `int` | The green component of the color to set (0-255). |
| `Blue` | `int` | The blue component of the color to set (0-255). |

**Returns: `any`**
#### Example
```sql
SELECT
	CASE JobTitle
		WHEN 'Chief Executive Officer' THEN SetFontColor(__address, 255,0,0)
		WHEN 'Vice President of Engineering' THEN SetFontColor(__address, 255,255,0)
		WHEN 'Research and Development Manager' THEN SetFontColor(__address, 0,255,255)
		ELSE SetFontColor(__address, 0,0,0)
	END
FROM
	Employee
```


### `ClearBackgroundColor`
Clears the interior color of the cell with the specified address.

| Name | Type  | Description  |
|---|---|---|
| `Address` | `string` | Address of the cell whose interior color is to be cleared. Most likely will be the hidden column field that each table in QueryStorm has. Can be named range, qualified ('e.g. Sheet1!A4') or unqualified (e.g. 'A4') cell address. |

**Returns: `string`**
#### Example
```sql
SELECT
	CASE JobTitle
		WHEN 'Chief Executive Officer' THEN SetBackgroundColor(__address, 'red')
		WHEN 'Vice President of Engineering' THEN SetBackgroundColor(__address, 'yellow')
		WHEN 'Research and Development Manager' THEN SetBackgroundColor(__address, 'teal')
		ELSE ClearBackgroundColor(__address)
	END
FROM
	Employee
```


### `RefreshPivotTable`
Refreshes the table with the given name.

| Name | Type  | Description  |
|---|---|---|
| `PivotTableName` | `string` | The name of the pivot table to refresh. Note: pivot table names should be as unique as possible, as the function has no way of distinguishing two pivot tables with the same name in two different open workooks. |

**Returns: `string`**
#### Example
```sql
SELECT RefreshPivotTable('PivotTable_Employees_rpt123');
```


### `Vba`
The way to implement user defined functions in Live mode. Calls the specified VBA function or subroutine passing the specified parameters and returns any results back. Useful for running custom calculations, but also for custom actions like sending emails, writing to files, and formatting cells. NOTE: For a function to return a value it MUST BE INSIDE A VBA MODULE (as opposed to being inside a sheet or the workbook), otherwise the return value will be Null.

| Name | Type  | Description  |
|---|---|---|
| `FunctionOrSub` | `string` | Name of the VBA function or subroutine. If the desired function/subroutine is inside a module the name doesn't need to be qualified, otherwise it does (e.g. 'Sheet1.SubABC' or 'ThisWorkbook.SubABC'). |
| `Arg1...ArgN` | `any` | Argument(s) to pass to the specified vba function or subroutine. |

**Returns: `any`**
#### Example
```sql
--NOTE: For a function to return a value it MUST BE INSIDE A VBA MODULE (as opposed to being inside a sheet or the workbook). 

-- For each person, calculate the sum of the digits in their IdNumber (because why not?).
-- Assumes the function 'SumAllDigits' exists in an available module and that it can accept IdNumber as parameter.
SELECT vba('SumAllDigits', IdNumber) FROM Person; 
```


