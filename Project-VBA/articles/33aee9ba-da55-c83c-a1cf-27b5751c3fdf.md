
# Application.SetLTRTable Method (Project)
Sets column order from left to right, for a selected table in a report.

 **Last modified:** July 28, 2015


## Syntax

 _expression_. **SetLTRTable**

 _expression_A variable that represents an  **Application** object.


### Return value

 **Boolean**

 **True** if the column order is set from left to right; otherwise, **False**.


## Remarks

The  **SetLTRTable** method can be used to change the table columns from right-to-left order for languages such as Arabic, to left-to-right for languages such as English, German, and French.

If a report is not active, the  **SetLTRTable** method displays a dialog box with run-time error 1100, "The method is not available in this situation."


## See also


#### Concepts


 [Application Object](8eb91712-7784-a102-38c0-19bb056c27e9.md)
#### Other resources


 [SetRTLTable](92dc18e3-fa84-a4b2-d032-aa32a4e3957d.md)
 [ReportTable Object](db9846c7-fd53-ae5a-7a43-35dfc60f4fe4.md)
 [Shape.Table Property](5e1fc97f-8ac8-db26-3a2d-c39261c23588.md)
