---
title: Window.SplitColumn Property (Excel)
keywords: vbaxl10.chm356112
f1_keywords:
- vbaxl10.chm356112
ms.prod: excel
api_name:
- Excel.Window.SplitColumn
ms.assetid: 699e2919-8786-4616-2363-78c3e01e4875
ms.date: 06/08/2017
---


# Window.SplitColumn Property (Excel)

Returns or sets the column number where the window is split into panes (the number of columns to the left of the split line). Read/write  **Long** .


## Syntax

 _expression_ . **SplitColumn**

 _expression_ A variable that represents a **Window** object.


## Example

This example splits the window and leaves 1.5 columns to the left of the split line.


```vb
Workbooks("BOOK1.XLS").Worksheets("Sheet1").Activate 
ActiveWindow.SplitColumn = 1.5
```


## See also


#### Concepts


[Window Object](window-object-excel.md)

