---
title: PageSetup.PrintTitleRows Property (Excel)
keywords: vbaxl10.chm473098
f1_keywords:
- vbaxl10.chm473098
ms.prod: excel
api_name:
- Excel.PageSetup.PrintTitleRows
ms.assetid: de01a1a9-a6f5-9eb4-5785-2993475c1a02
ms.date: 06/08/2017
---


# PageSetup.PrintTitleRows Property (Excel)

Returns or sets the rows that contain the cells to be repeated at the top of each page, as a string in A1-style notation in the language of the macro. Read/write  **String** .


## Syntax

 _expression_ . **PrintTitleRows**

 _expression_ A variable that represents a **PageSetup** object.


## Remarks

If you specify only part of a row or rows, Microsoft Excel expands the range to full rows.

Set this property to  **False** or to the empty string ("") to turn off title rows.

This property applies only to worksheet pages.


## Example

This example defines row three as the title row, and it defines columns one through three as the title columns.


```vb
Worksheets("Sheet1").Activate 
ActiveSheet.PageSetup.PrintTitleRows = ActiveSheet.Rows(3).Address 
ActiveSheet.PageSetup.PrintTitleColumns = _ 
 ActiveSheet.Columns("A:C").Address
```


## See also


#### Concepts


[PageSetup Object](pagesetup-object-excel.md)

