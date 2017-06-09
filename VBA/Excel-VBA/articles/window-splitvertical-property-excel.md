---
title: Window.SplitVertical Property (Excel)
keywords: vbaxl10.chm356115
f1_keywords:
- vbaxl10.chm356115
ms.prod: excel
api_name:
- Excel.Window.SplitVertical
ms.assetid: 2e683391-b5c3-0d4d-94a3-0afe82e3965a
ms.date: 06/08/2017
---


# Window.SplitVertical Property (Excel)

Returns or sets the location of the vertical window split, in points. Read/write  **Double** .


## Syntax

 _expression_ . **SplitVertical**

 _expression_ A variable that represents a **Window** object.


## Example

This example sets the vertical split for the active window to 216 points (3 inches).


```vb
Workbooks("BOOK1.XLS").Worksheets("Sheet1").Activate 
ActiveWindow.SplitVertical = 216
```


## See also


#### Concepts


[Window Object](window-object-excel.md)

