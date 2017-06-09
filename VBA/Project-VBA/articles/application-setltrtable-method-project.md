---
title: Application.SetLTRTable Method (Project)
keywords: vbapj.chm1520
f1_keywords:
- vbapj.chm1520
ms.prod: project-server
ms.assetid: 33aee9ba-da55-c83c-a1cf-27b5751c3fdf
ms.date: 06/08/2017
---


# Application.SetLTRTable Method (Project)
Sets column order from left to right, for a selected table in a report.

## Syntax

 _expression_. **SetLTRTable**

 _expression_ A variable that represents an **Application** object.


### Return value

 **Boolean**

 **True** if the column order is set from left to right; otherwise, **False**.


## Remarks

The  **SetLTRTable** method can be used to change the table columns from right-to-left order for languages such as Arabic, to left-to-right for languages such as English, German, and French.

If a report is not active, the  **SetLTRTable** method displays a dialog box with run-time error 1100, "The method is not available in this situation."


## See also


#### Concepts


[Application Object](application-object-project.md)
#### Other resources


[SetRTLTable](application-setrtltable-method-project.md)
[ReportTable Object](reporttable-object-project.md)
[Shape.Table Property](shape-table-property-project.md)
