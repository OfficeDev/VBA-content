---
title: Window.DisplayZeros Property (Excel)
keywords: vbaxl10.chm356090
f1_keywords:
- vbaxl10.chm356090
ms.prod: excel
api_name:
- Excel.Window.DisplayZeros
ms.assetid: cddb671b-5b7f-c2a8-1527-bfe0bfdced78
ms.date: 06/08/2017
---


# Window.DisplayZeros Property (Excel)

 **True** if zero values are displayed. Read/write **Boolean** .


## Syntax

 _expression_ . **DisplayZeros**

 _expression_ A variable that represents a **Window** object.


## Remarks

This property applies only to worksheets and macro sheets.


## Example

This example sets the active window in Book1.xls to display zero values.


```vb
Workbooks("BOOK1.XLS").Worksheets("Sheet1").Activate 
ActiveWindow.DisplayZeros = True 

```


## See also


#### Concepts


[Window Object](window-object-excel.md)

