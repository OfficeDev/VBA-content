---
title: Worksheet.Columns Property (Excel)
keywords: vbaxl10.chm175086
f1_keywords:
- vbaxl10.chm175086
ms.prod: excel
api_name:
- Excel.Worksheet.Columns
ms.assetid: 41c18561-2a87-b975-e212-97f39fe10393
ms.date: 06/08/2017
---


# Worksheet.Columns Property (Excel)

Returns a  **[Range](range-object-excel.md)** object that represents all the columns on the specified worksheet.


## Syntax

 _expression_ . **Columns**

 _expression_ A variable that represents a **Worksheet** object.


## Remarks

`Columns` without an object qualifier is equivalent to `ActiveSheet.Columns`. If the active sheet isn't a worksheet, then `Columns` fails.

To return a single column, include an index in parentheses. For example, `Columns(1)` and `Columns("A")` return the first column.


## Example

This example formats the font of column one (column A) on Sheet1 as bold.


```vb
Worksheets("Sheet1").Columns(1).Font.Bold = True
```


## See also


#### Related

[Range.Columns Property](range-columns-property-excel.md)


#### Concepts

[Worksheet Object](worksheet-object-excel.md)

