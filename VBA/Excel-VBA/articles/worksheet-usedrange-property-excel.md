---
title: Worksheet.UsedRange Property (Excel)
keywords: vbaxl10.chm175134
f1_keywords:
- vbaxl10.chm175134
ms.prod: excel
api_name:
- Excel.Worksheet.UsedRange
ms.assetid: f004b93c-d785-de19-1fb4-bbe0b2e9b6cd
ms.date: 06/08/2017
---


# Worksheet.UsedRange Property (Excel)

Returns a  **[Range](range-object-excel.md)** object that represents the used range on the specified worksheet. Read-only.


## Syntax

 _expression_ . **UsedRange**

 _expression_ A variable that represents a **Worksheet** object.


## Example

This example selects the used range on Sheet1.


```vb
Worksheets("Sheet1").Activate 
ActiveSheet.UsedRange.Select
```


## See also


#### Concepts


[Worksheet Object](worksheet-object-excel.md)

