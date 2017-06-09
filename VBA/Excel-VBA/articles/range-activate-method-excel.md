---
title: Range.Activate Method (Excel)
keywords: vbaxl10.chm144074
f1_keywords:
- vbaxl10.chm144074
ms.prod: excel
api_name:
- Excel.Range.Activate
ms.assetid: a0050055-84e7-7611-a961-887fcb063369
ms.date: 06/08/2017
---


# Range.Activate Method (Excel)

Activates a single cell, which must be inside the current selection. To select a range of cells, use the  **[Select](range-select-method-excel.md)** method.


## Syntax

 _expression_ . **Activate**

 _expression_ A variable that represents a **Range** object.


### Return Value

Variant


## Example

This example selects cells A1:C3 on Sheet1 and then makes cell B2 the active cell.


```vb
Worksheets("Sheet1").Activate 
Range("A1:C3").Select 
Range("B2").Activate
```


## See also


#### Concepts


[Range Object](range-object-excel.md)

