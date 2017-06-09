---
title: Worksheet.CircularReference Property (Excel)
keywords: vbaxl10.chm175084
f1_keywords:
- vbaxl10.chm175084
ms.prod: excel
api_name:
- Excel.Worksheet.CircularReference
ms.assetid: 422c447d-a964-c17c-bb43-14254f962a89
ms.date: 06/08/2017
---


# Worksheet.CircularReference Property (Excel)

Returns a  **[Range](range-object-excel.md)** object that represents the range containing the first circular reference on the sheet, or returns **Nothing** if there's no circular reference on the sheet. The circular reference must be removed before calculation can proceed.


## Syntax

 _expression_ . **CircularReference**

 _expression_ A variable that represents a **Worksheet** object.


## Example

This example selects the first cell in the first circular reference on Sheet1.


```vb
Worksheets("Sheet1").CircularReference.Select
```


## See also


#### Concepts


[Worksheet Object](worksheet-object-excel.md)

