# Automation jobs

Jobs enable automatic execution of embedded queries in response to certain events. This is useful for creating refreshable reports as well as for adding application-like behavior to workbooks.

In the example below, a C# script that fetches data from a web service is triggered by a button on the sheet. The script updates the data in a table, which in turn updates the graph. The graph is then controlled by a slicer.
 
![Automation with C#](https://www.querystorm.com/Downloads/Images/exappgif.gif "Automation with C#")

Each job consists of a set of triggers and actions. 

![Automation job](../images/job.png "Automation job")

Triggers can be timers or user actions (e.g. cell value changed, button clicked...). When ever **any** trigger in a job *fires*, all the queries/scripts in that job are executed in sequence. The order of execution can be adjusted using the *Up*/*Down* buttons in the *Automation* pane. 

If queries do not depend on each other, they can be executed in parallel by separate jobs. If two jobs define the same trigger (e.g. cell *A1* changed), when that even occurs, both jobs will execute in parallel.

You can set up jobs in the Automation pane. There are also handy shortcuts in the context menus in Excel. The context menu for ActiveX buttons is shown below. There are also context menus for ranges and tables. 

![Automation context menus](../images/automation_context_menus.png "Automation context menus") 
