---
title: Window.WindowNumber Property (Excel)
keywords: vbaxl10.chm356124
f1_keywords:
- vbaxl10.chm356124
ms.prod: excel
api_name:
- Excel.Window.WindowNumber
ms.assetid: 42dc6fa2-8c10-41d8-2f74-95401e154094
ms.date: 06/08/2017
---


# Window.WindowNumber Property (Excel)

Returns the window number. For example, a window named "Book1.xls:2" has 2 as its window number. Most windows have the window number 1. Read-only  **Long** .


## Syntax

 _expression_ . **WindowNumber**

 _expression_ A variable that represents a **Window** object.


## Remarks

The window number isn't the same as the window index (the return value of the  **Index** property), which is the position of the window within the **Windows** collection.


## Example

This example creates a new window of the active window and then displays the window number of the new window.


```vb
ActiveWindow.NewWindow 
MsgBox ActiveWindow.WindowNumber
```


## See also


#### Concepts


[Window Object](window-object-excel.md)

