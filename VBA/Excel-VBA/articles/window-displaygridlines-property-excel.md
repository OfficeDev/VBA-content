---
title: Window.DisplayGridlines Property (Excel)
keywords: vbaxl10.chm356083
f1_keywords:
- vbaxl10.chm356083
ms.prod: excel
api_name:
- Excel.Window.DisplayGridlines
ms.assetid: d4253c7f-bed2-6e58-9b04-479355f70561
ms.date: 06/08/2017
---


# Window.DisplayGridlines Property (Excel)

 **True** if gridlines are displayed. Read/write **Boolean** .


## Syntax

 _expression_ . **DisplayGridlines**

 _expression_ A variable that represents a **Window** object.


## Remarks

This property applies only to worksheets and macro sheets.

This property affects only displayed gridlines. Use the  **[PrintGridlines](pagesetup-printgridlines-property-excel.md)** property to control the printing of gridlines.


## Example

This example toggles the display of gridlines in the active window in Book1.xls.


```vb
Workbooks("BOOK1.XLS").Worksheets("Sheet1").Activate 
ActiveWindow.DisplayGridlines = Not(ActiveWindow.DisplayGridlines) 

```


## See also


#### Concepts


[Window Object](window-object-excel.md)

