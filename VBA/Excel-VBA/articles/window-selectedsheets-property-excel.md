---
title: Window.SelectedSheets Property (Excel)
keywords: vbaxl10.chm356108
f1_keywords:
- vbaxl10.chm356108
ms.prod: excel
api_name:
- Excel.Window.SelectedSheets
ms.assetid: 3be18be3-895b-131b-7416-270536b84e23
ms.date: 06/08/2017
---


# Window.SelectedSheets Property (Excel)

Returns a  **[Sheets](sheets-object-excel.md)** collection that represents all the selected sheets in the specified window. Read-only.


## Syntax

 _expression_ . **SelectedSheets**

 _expression_ A variable that represents a **Window** object.


## Example

This example displays a message if Sheet1 is selected in Book1.xls.


```vb
For Each sh In Workbooks("BOOK1.XLS").Windows(1).SelectedSheets 
 If sh.Name = "Sheet1" Then 
 MsgBox "Sheet1 is selected" 
 Exit For 
 End If 
Next
```


## See also


#### Concepts


[Window Object](window-object-excel.md)

