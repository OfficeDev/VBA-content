---
title: Application.Calculate Method (Excel)
keywords: vbaxl10.chm183084
f1_keywords:
- vbaxl10.chm183084
ms.prod: excel
api_name:
- Excel.Application.Calculate
ms.assetid: 2818a08b-1c02-9f10-db03-db509a251f60
ms.date: 06/08/2017
---


# Application.Calculate Method (Excel)

Calculates all open workbooks, a specific worksheet in a workbook, or a specified range of cells on a worksheet, as shown in the following table.


## Syntax

 _expression_ . **Calculate**

 _expression_ A variable that represents an **Application** object.


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


[Application Object](application-object-excel.md)

