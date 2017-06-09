---
title: Application.SetRTLTable Method (Project)
keywords: vbapj.chm1519
f1_keywords:
- vbapj.chm1519
ms.prod: project-server
ms.assetid: 92dc18e3-fa84-a4b2-d032-aa32a4e3957d
ms.date: 06/08/2017
---


# Application.SetRTLTable Method (Project)
Sets column order from right to left, for a selected table in a report.

## Syntax

 _expression_. **SetRTLTable**

 _expression_ A variable that represents a **Application** object.


### Return value

 **Boolean**

 **True** if the column order is set from right to left; otherwise, **False**.


## Remarks

The  **SetRTLTable** method can be used to change the table columns from left-to-right order for languages such as English, to right-to-left for languages such as Arabic, Farsi, and Hebrew.

If a report is not active, the  **SetRTLTable** method displays a dialog box with run-time error 1100, "The method is not available in this situation."


## See also


#### Concepts


[Application Object](application-object-project.md)
#### Other resources


[SetLTRTable](application-setltrtable-method-project.md)
[ReportTable Object](reporttable-object-project.md)
[Shape.Table Property](shape-table-property-project.md)
