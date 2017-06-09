---
title: Window.Top Property (Excel)
keywords: vbaxl10.chm356117
f1_keywords:
- vbaxl10.chm356117
ms.prod: excel
api_name:
- Excel.Window.Top
ms.assetid: e04d6641-9788-1e0b-9343-011c414c31fc
ms.date: 06/08/2017
---


# Window.Top Property (Excel)

Returns or sets a  **Double** value that represents the distance, in points, from the top edge of the window to the top edge of the usable area (below the menus, any toolbars docked at the top, and the formula bar).


## Syntax

 _expression_ . **Top**

 _expression_ A variable that represents a **Window** object.


## Remarks

You cannot set this property for a maximized window. Use the  **[WindowState](window-windowstate-property-excel.md)** property to return or set the state of the window.


## Example

This example arranges windows one and two horizontally; in other words, each window occupies half the available vertical space and all the available horizontal space in the application window's client area. For this example to work, there must be only two worksheet windows open.


```vb
Windows.Arrange xlArrangeTiled 
ah = Windows(1).Height ' available height 
aw = Windows(1).Width + Windows(2).Width ' available width 
With Windows(1) 
 .Width = aw 
 .Height = ah / 2 
 .Left = 0 
End With 
With Windows(2) 
 .Width = aw 
 .Height = ah / 2 
 .Top = ah / 2 
 .Left = 0 
End With
```


## See also


#### Concepts


[Window Object](window-object-excel.md)

