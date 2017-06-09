---
title: Window.ScrollRow Property (Excel)
keywords: vbaxl10.chm356106
f1_keywords:
- vbaxl10.chm356106
ms.prod: excel
api_name:
- Excel.Window.ScrollRow
ms.assetid: 5fd21ea8-a173-e502-042d-57903bcd43e5
ms.date: 06/08/2017
---


# Window.ScrollRow Property (Excel)

Returns or sets the number of the row that appears at the top of the pane or window. Read/write  **Long** .


## Syntax

 _expression_ . **ScrollRow**

 _expression_ A variable that represents a **Window** object.


## Remarks

If the window is split, the  **ScrollRow** property of the **[Window](window-object-excel.md)** object refers to the upper-left pane. If the panes are frozen, the **ScrollRow** property of the **Window** object excludes the frozen areas.


## Example

This example moves row ten to the top of the window.


```vb
Worksheets("Sheet1").Activate 
ActiveWindow.ScrollRow = 10
```


## See also


#### Concepts


[Window Object](window-object-excel.md)

