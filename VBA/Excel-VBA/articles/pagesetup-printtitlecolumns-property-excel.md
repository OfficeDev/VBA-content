---
title: PageSetup.PrintTitleColumns Property (Excel)
keywords: vbaxl10.chm473097
f1_keywords:
- vbaxl10.chm473097
ms.prod: excel
api_name:
- Excel.PageSetup.PrintTitleColumns
ms.assetid: 860cf212-0fbb-f3ec-c9ce-a0df57b39b7f
ms.date: 06/08/2017
---


# PageSetup.PrintTitleColumns Property (Excel)

Returns or sets the columns that contain the cells to be repeated on the left side of each page, as a string in A1-style notation in the language of the macro. Read/write  **String** .


## Syntax

 _expression_ . **PrintTitleColumns**

 _expression_ A variable that represents a **PageSetup** object.


## Remarks

If you specify only part of a column or columns, Microsoft Excel expands the range to full columns.

Set this property to  **False** or to the empty string ("") to turn off title columns.

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

