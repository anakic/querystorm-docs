# Under the hood
You might be curious how things work under the hood. If you're reading this, you're likely to be fairly technical, after all. Here's a rough overview of how the C# engine works...

## Dynamically generated types
For a given Excel table, QueryStorm will dynamically generate a type (using `Reflection.Emit`) that will represent a table's row, in order to expose columns as strongly typed properties. This is done for each table in scope. Row objects don't have a copy of the row data, they only contain the index of the row in the table. Each property *knows* which column index it refers to, so it uses the row index and column index to access the correct value from the table's cache. This makes row objects extremely lightweight (they only contain a single integer).

## The cache
So what's the table cache? QueryStorm internally maintains an 2d-array (`object[,]`) with the data for each table. This array is the table cache. The cache is selectively updated as changes are made to the Excel table (e.g. user changes a value). If the cache becomes corrupt for any reason, the user can explicitly reload it from the context menu of the table in the object explorer (this shouldn't really happen but...). Anyway, since the cache is entirely in memory, queries are very fast; e.g. ~30ms for reading 100k rows with 10 columns. Updating a row's property writes the new value to the cache, and the cache keeps track of which regions have been changed. When the user calls `SaveChanges()`, a minimal block of the cache (that contains all the changes) is written into the Excel table, making updates very fast as well.

## Roslyn
The execution of the scripts as well as the IDE features (intellisense, error squiggles, code auto-formatting) are powered by the Roslyn compiler. Each time the script text changes (user types a character), the code is analyzed and the editor updated to display any errors, intellisense suggestions, function documentation etc... Upon execution, a new thread is started in order to not block the UI and to enable the user to cancel execution. The version of C# that is made available to the user is (currently) C# 7, but in general will be updated to the latest version when ever a new version comes out.

## Planned improvements
- C# debugger
- Ability to store keep referenced dll's inside the workbook
- Ability to reference NuGet packages
- Editor features: rename and look up symbol, include namespace using
- An `IQueryable` provider for indexing. The indexing mechanism already exists (built for QueryStorm's SQLite engine), but it isn't yet being used by the C# engine.

## Acknowledgments
The C# scripting engine in QueryStorm uses the Roslyn compiler for execution and IDE support. Roslyn has a bit of a learning curve, but these excellent resources made things a lot easier:

- [Roslynpad](https://github.com/aelij/RoslynPad "RoslynPad") by Aelij
- Josh Varty's [Roslyn series](https://joshvarty.wordpress.com/learn-roslyn-now/ "Josh Varty's Roslyn blog")
- [Roslyn docs](https://github.com/dotnet/roslyn "Roslyn docs") on Github

QueryStorm also uses numerous other open source libraries, most notably [AvalonEdit](http://avalonedit.net/ "AvalonEdit") and [SQLite](http://sqlite.org/ "SQLite").