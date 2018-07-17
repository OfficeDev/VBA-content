---
title: Window.DisplayFormulas Property (Excel)
keywords: vbaxl10.chm356082
f1_keywords:
- vbaxl10.chm356082
ms.prod: excel
api_name:
- Excel.Window.DisplayFormulas
ms.assetid: 04e75e40-4eb9-93f9-73b2-4024a1c1151d
ms.date: 06/08/2017
---


# Window.DisplayFormulas Property (Excel)

 **True** if the window is displaying formulas; **False** if the window is displaying values. Read/write **Boolean** .


## Syntax

 _expression_ . **DisplayFormulas**

 _expression_ A variable that represents a **Window** object.


## Remarks

This property applies only to worksheets and macro sheets.


## Example

This example changes the active window in Book1.xls to display formulas.


```vb
Workbooks("BOOK1.XLS").Worksheets("Sheet1").Activate 
ActiveWindow.DisplayFormulas = True 

```


## See also


#### Concepts


[Window Object](window-object-excel.md)

