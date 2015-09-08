
# Screen.ActiveReport Property (Access)

 **Last modified:** July 28, 2015

You can use the  **ActiveReport** property together with the ** [Screen](00743775-071b-9ccd-7687-f3b992e9346e.md)**object to identify or refer to the report that has the focus. Read-only  **Report** object.

## Syntax

 _expression_. **ActiveReport**

 _expression_A variable that represents a  **Screen** object.


## Remarks

This property setting contains a reference to the  ** [Report](6f77c1b4-a9ce-7caa-204c-fe0755c6f9df.md)**object that has the focus at run time.

You can use the  **ActiveReport** property to refer to an active report together with one of its properties or methods. The following example displays the **Name**property setting of the active report.




```
Dim rptCurrentReport As Report 
Set rptCurrentReport = Screen.ActiveReport 
MsgBox "Current report is " &amp; rptCurrentReport.Name
```

If no report has the focus when you use the  **ActiveReport** property, an error occurs.


## See also


#### Concepts


 [Screen Object](00743775-071b-9ccd-7687-f3b992e9346e.md)
#### Other resources


 [Screen Object Members](82c9e4cb-95a9-6842-2629-bcd71c81838f.md)
