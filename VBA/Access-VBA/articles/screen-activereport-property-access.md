---
title: Screen.ActiveReport Property (Access)
keywords: vbaac10.chm12491
f1_keywords:
- vbaac10.chm12491
ms.prod: access
api_name:
- Access.Screen.ActiveReport
ms.assetid: efcf6bfd-2749-5b5c-d7ca-a26168bfcb65
ms.date: 06/08/2017
---


# Screen.ActiveReport Property (Access)

You can use the  **ActiveReport** property together with the **[Screen](screen-object-access.md)** object to identify or refer to the report that has the focus. Read-only **Report** object.


## Syntax

 _expression_. **ActiveReport**

 _expression_ A variable that represents a **Screen** object.


## Remarks

This property setting contains a reference to the  **[Report](report-object-access.md)** object that has the focus at run time.

You can use the  **ActiveReport** property to refer to an active report together with one of its properties or methods. The following example displays the **Name** property setting of the active report.




```vb
Dim rptCurrentReport As Report 
Set rptCurrentReport = Screen.ActiveReport 
MsgBox "Current report is " &; rptCurrentReport.Name
```

If no report has the focus when you use the  **ActiveReport** property, an error occurs.


## See also


#### Concepts


[Screen Object](screen-object-access.md)

