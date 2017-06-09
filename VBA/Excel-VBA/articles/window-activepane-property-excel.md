---
title: Window.ActivePane Property (Excel)
keywords: vbaxl10.chm356078
f1_keywords:
- vbaxl10.chm356078
ms.prod: excel
api_name:
- Excel.Window.ActivePane
ms.assetid: f518802d-8624-6e61-d76a-d318149e0142
ms.date: 06/08/2017
---


# Window.ActivePane Property (Excel)

Returns a  **[Pane](pane-object-excel.md)** object that represents the active pane in the window. Read-only.


## Syntax

 _expression_ . **ActivePane**

 _expression_ A variable that represents a **Window** object.


## Remarks

This property can be used only on worksheets and macro sheets.

This property returns a  **Pane** object. You must use the **[Index](pane-index-property-excel.md)** property to obtain the index of the active pane.


## Example

This example activates the next pane of the active window in Book1.xls. You cannot activate the next pane if the panes are frozen. The example must be run from a workbook other than Book1.xls. Before running the example, make sure that Book1.xls has either two or four panes in the active worksheet.


```vb
Workbooks("BOOK1.XLS").Activate 
If not ActiveWindow.FreezePanes Then 
 With ActiveWindow 
 i = .ActivePane.Index 
 If i = .Panes.Count Then 
 .Panes(1).Activate 
 Else 
 .Panes(i+1).Activate 
 End If 
 End With 
End If
```


## See also


#### Concepts


[Window Object](window-object-excel.md)

