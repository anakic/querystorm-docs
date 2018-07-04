# Exchange rates visualizer

This example demonstrates loading exchange rates (from a web API, via C#) for a given currency for the past 36 months and displaying them interactively in a graph.

!!! Download
	Download the workbook from [here](../demofiles/rates.xlsx)

![Automation with C#](https://www.querystorm.com/Downloads/Images/exappgif.gif "Automation with C#")

Once you enter the base currency and press *Refresh*, the data for the past 36 months will be loaded from the web api and pushed into an Excel table in *Sheet2*. After that, the c# script refreshes the pivot table which, in turn, refreshes the graph in *Sheet1*. The slicer is then used to control what gets displayed on the graph. 

## How it works
The exchange rates data is loaded by a C# function that fetches it from a web API. For the purposes of the demo, the C# function is then called from a SQLite query, which is in turn called by a C# script that is triggered by the *Refresh* button. This is done solely for the purposes of demonstrating the C#/SQL interactivity. 

Check out the embedded code and automation in the [demo workbook](../demofiles/rates.xlsx) to see how it works.

![Under the hood](../images/rates_internals.png)

## Summary

C# code was used for loading data from a web service, while the workbook was used for storage and the user interface. The combination of application and document makes it remarkably easy to build, and the resulting file requires no additional infrastructure (web hosting, database server).  
