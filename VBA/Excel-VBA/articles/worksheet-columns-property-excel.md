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

Returns a  **[Range](range-object-excel.md)** object that represents all the columns on the active worksheet. If the active document isn't a worksheet, the **Columns** property fails.


## Syntax

 _expression_ . **Columns**

 _expression_ A variable that represents a **Worksheet** object.


## Remarks

Using this property without an object qualifier is equivalent to using  `ActiveSheet.Columns`.

When applied to a  **Range** object that's a multiple-area selection, this property returns columns from only the first area of the range. For example, if the **Range** object has two areas — A1:B2 and C3:D4 — `Selection.Columns.Count` returns 2, not 4. To use this property on a range that may contain a multiple-area selection, test `Areas.Count` to determine whether the range contains more than one area. If it does, loop over each area in the range.


## Example

This example formats the font of column one (column A) on Sheet1 as bold.


```vb
Worksheets("Sheet1").Columns(1).Font.Bold = True
```


## See also


#### Concepts


[Worksheet Object](worksheet-object-excel.md)

