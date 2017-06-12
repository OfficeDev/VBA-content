---
title: Application.ShowStartupDialog Property (Excel)
keywords: vbaxl10.chm133287
f1_keywords:
- vbaxl10.chm133287
ms.prod: excel
api_name:
- Excel.Application.ShowStartupDialog
ms.assetid: 8ea751c4-a4b1-a84a-9566-c4de8c5b9f67
ms.date: 06/08/2017
---


# Application.ShowStartupDialog Property (Excel)

Returns  **True** (default is **False** ) when the New Workbook task pane appears for a Microsoft Excel application. Read/write **Boolean** .


## Syntax

 _expression_ . **ShowStartupDialog**

 _expression_ A variable that represents an **Application** object.


## Example

In this example, Microsoft Excel determines if the New Workbook task pane appears and notifies the user.


```vb
Sub CheckStartupDialog() 
 
 ' Determine if the New Workbook task pane is enabled. 
 If Application.ShowStartupDialog = False Then 
 MsgBox "ShowStartupDialog is set to False." 
 Else 
 MsgBox "ShowStartupDialog is set to True." 
 End If 
 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

