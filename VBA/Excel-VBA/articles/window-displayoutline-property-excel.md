---
title: Window.DisplayOutline Property (Excel)
keywords: vbaxl10.chm356086
f1_keywords:
- vbaxl10.chm356086
ms.prod: excel
api_name:
- Excel.Window.DisplayOutline
ms.assetid: 3934e907-1792-6ff3-6529-dd1dd45ce221
ms.date: 06/08/2017
---


# Window.DisplayOutline Property (Excel)

 **True** if outline symbols are displayed. Read/write **Boolean** .


## Syntax

 _expression_ . **DisplayOutline**

 _expression_ A variable that represents a **Window** object.


## Remarks

This property applies only to worksheets and macro sheets.


## Example

This example displays outline symbols for the active window in Book1.xls.


```vb
Workbooks("BOOK1.XLS").Worksheets("Sheet1").Activate 
ActiveWindow.DisplayOutline = True 

```


## See also


#### Concepts


[Window Object](window-object-excel.md)

