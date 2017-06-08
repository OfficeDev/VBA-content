---
title: Window.FreezePanes Property (Excel)
keywords: vbaxl10.chm356092
f1_keywords:
- vbaxl10.chm356092
ms.prod: excel
api_name:
- Excel.Window.FreezePanes
ms.assetid: fd8c7b3b-4f70-72bd-68e4-a34442192a4e
ms.date: 06/08/2017
---


# Window.FreezePanes Property (Excel)

 **True** if split panes are frozen. Read/write **Boolean** .


## Syntax

 _expression_ . **FreezePanes**

 _expression_ A variable that represents a **Window** object.


## Remarks

It's possible for  **FreezePanes** to be **True** and **[Split](window-split-property-excel.md)** to be **False** , or vice versa.

This property applies only to worksheets and macro sheets.


## Example

This example freezes split panes in the active window in Book1.xls.


```vb
Workbooks("BOOK1.XLS").Worksheets("Sheet1").Activate 
ActiveWindow.FreezePanes = True
```


## See also


#### Concepts


[Window Object](window-object-excel.md)

