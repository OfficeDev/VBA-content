---
title: Worksheet.StandardWidth Property (Excel)
keywords: vbaxl10.chm175130
f1_keywords:
- vbaxl10.chm175130
ms.prod: excel
api_name:
- Excel.Worksheet.StandardWidth
ms.assetid: 6792ce79-0a73-fcbd-ea52-7d7aee7b9932
ms.date: 06/08/2017
---


# Worksheet.StandardWidth Property (Excel)

Returns or sets the standard (default) width of all the columns in the worksheet. Read/write  **Double** .


## Syntax

 _expression_ . **StandardWidth**

 _expression_ A variable that represents a **Worksheet** object.


## Remarks

One unit of column width is equal to the width of one character in the Normal style. For proportional fonts, the width of the character 0 (zero) is used.


## Example

This example sets the width of column one on Sheet1 to the standard width.


```vb
Worksheets("Sheet1").Columns(1).ColumnWidth = _ 
 Worksheets("Sheet1").StandardWidth
```


## See also


#### Concepts


[Worksheet Object](worksheet-object-excel.md)

