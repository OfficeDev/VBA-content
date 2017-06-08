---
title: Worksheet.Calculate Method (Excel)
keywords: vbaxl10.chm175078
f1_keywords:
- vbaxl10.chm175078
ms.prod: excel
api_name:
- Excel.Worksheet.Calculate
ms.assetid: 7e807ae0-cd97-d95b-f4c4-af1e5674943e
ms.date: 06/08/2017
---


# Worksheet.Calculate Method (Excel)

Calculates all open workbooks, a specific worksheet in a workbook, or a specified range of cells on a worksheet, as shown in the following table.


## Syntax

 _expression_ . **Calculate**

 _expression_ A variable that represents a **Worksheet** object.


## Remarks





|**To calculate**|**Follow this example**|
|:-----|:-----|
|All open workbooks| `Application.Calculate` (or just `Calculate`)|
|A specific worksheet| `Worksheets(1).Calculate`|
|A specified range| `Worksheets(1).Rows(2).Calculate`|

## Example

This example calculates the formulas in columns A, B, and C in the used range on Sheet1.


```vb
Worksheets("Sheet1").UsedRange.Columns("A:C").Calculate
```


## See also


#### Concepts


[Worksheet Object](worksheet-object-excel.md)

