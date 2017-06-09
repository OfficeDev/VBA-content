---
title: Window.OnWindow Property (Excel)
keywords: vbaxl10.chm356100
f1_keywords:
- vbaxl10.chm356100
ms.prod: excel
api_name:
- Excel.Window.OnWindow
ms.assetid: 928415d0-075b-acea-ab47-5d971a9b86b6
ms.date: 06/08/2017
---


# Window.OnWindow Property (Excel)

Returns or sets the name of the procedure that's run whenever you activate a window. Read/write  **String** .


## Syntax

 _expression_ . **OnWindow**

 _expression_ A variable that represents a **Window** object.


## Remarks

The procedure specified by this property isn't run when other procedures switch to the window or when a command to switch to a window is received through a DDE channel. Instead, the procedure responds to the user's actions, such as clicking a window with the mouse.

If a worksheet or macro sheet has an Auto_Activate or Auto_Deactivate macro defined for it, those macros will be run after the procedure specified by the  **[OnWindow](application-onwindow-property-excel.md)** property.


## Example

This example causes the WindowActivate procedure to be run whenever window one is activated.


```vb
ThisWorkbook.Windows(1).OnWindow = "WindowActivate"
```


## See also


#### Concepts


[Window Object](window-object-excel.md)

