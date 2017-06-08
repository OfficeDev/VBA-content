---
title: Application.ActiveCell Property (Excel)
keywords: vbaxl10.chm183074
f1_keywords:
- vbaxl10.chm183074
ms.prod: excel
api_name:
- Excel.Application.ActiveCell
ms.assetid: 7ebfbec8-dc4e-36c5-188a-347d42649e76
ms.date: 06/08/2017
---


# Application.ActiveCell Property (Excel)

Returns a  **[Range](range-object-excel.md)** object that represents the active cell in the active window (the window on top) or in the specified window. If the window isn't displaying a worksheet, this property fails. Read-only.


## Syntax

 _expression_ . **ActiveCell**

 _expression_ A variable that represents an **Application** object.


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


[Application Object](application-object-excel.md)

