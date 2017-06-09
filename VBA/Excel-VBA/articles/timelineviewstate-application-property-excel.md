---
title: TimelineViewState.Application Property (Excel)
keywords: vbaxl10.chm951073
f1_keywords:
- vbaxl10.chm951073
ms.prod: excel
ms.assetid: b00518cc-b584-c562-0ae3-cf1e24844bdd
ms.date: 06/08/2017
---


# TimelineViewState.Application Property (Excel)

Returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. Read-only.


## Syntax

 _expression_ . **Application**

 _expression_ A variable that represents a[TimelineViewState Object (Excel)](timelineviewstate-object-excel.md) object.


## Example

This example displays a message about the application that created  `myObject`.


```vb
Set myObject = ActiveWorkbook 
If myObject.Application.Value = "Microsoft Excel" Then 
 MsgBox "This is an Excel Application object." 
Else 
 MsgBox "This is not an Excel Application object." 
End If
```


## Property value

 **APPLICATION**


## See also


#### Other resources



[TimelineViewState Object](timelineviewstate-object-excel.md)

