---
title: Range.DirectDependents Property (Excel)
keywords: vbaxl10.chm144118
f1_keywords:
- vbaxl10.chm144118
ms.prod: excel
api_name:
- Excel.Range.DirectDependents
ms.assetid: 266b054e-6838-ffe7-14e2-e712a149e5be
ms.date: 06/08/2017
---


# Range.DirectDependents Property (Excel)

Returns a  **[Range](range-object-excel.md)** object that represents the range containing all the direct dependents of a cell. This can be a multiple selection (a union of **Range** objects) if there's more than one dependent. Read-only **Range** object.


## Syntax

 _expression_ . **DirectDependents**

 _expression_ A variable that represents a **Range** object.


## Remarks

The  **Direct Dependents** property only works on the active sheet and can not trace remote references.


## Example

This example selects the direct dependents of cell A1 on Sheet1.


```vb
Worksheets("Sheet1").Activate 
Range("A1").DirectDependents.Select
```


## See also


#### Concepts


[Range Object](range-object-excel.md)

