---
title: Window.ActiveCell Property (Excel)
keywords: vbaxl10.chm356076
f1_keywords:
- vbaxl10.chm356076
ms.prod: excel
api_name:
- Excel.Window.ActiveCell
ms.assetid: 07ae9613-94b4-b3b9-c645-8acdabfebe86
ms.date: 06/08/2017
---


# Window.ActiveCell Property (Excel)

Returns a  **[Range](range-object-excel.md)** object that represents the active cell in the active window (the window on top) or in the specified window. If the window isn't displaying a worksheet, this property fails. Read-only.


## Syntax

 _expression_ . **ActiveCell**

 _expression_ A variable that represents a **Window** object.


## Remarks

If you don't specify an object qualifier, this property returns the active cell in the active window.

Be careful to distinguish between the active cell and the selection. The active cell is a single cell inside the current selection. The selection may contain more than one cell, but only one is the active cell.

The following expressions all return the active cell, and are all equivalent.




```vb
ActiveCell 
Application.ActiveCell 
ActiveWindow.ActiveCell 
Application.ActiveWindow.ActiveCell
```


## Example

This example uses a message box to display the value in the active cell. Because the  **ActiveCell** property fails if the active sheet isn't a worksheet, the example activates Sheet1 before using the **ActiveCell** property.


```vb
Worksheets("Sheet1").Activate 
MsgBox ActiveCell.Value
```

This example changes the font formatting for the active cell.




```vb
Worksheets("Sheet1").Activate 
With ActiveCell.Font 
 .Bold = True 
 .Italic = True 
End With
```


## See also


#### Concepts


[Window Object](window-object-excel.md)

