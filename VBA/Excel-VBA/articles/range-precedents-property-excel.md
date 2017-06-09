---
title: Range.Precedents Property (Excel)
keywords: vbaxl10.chm144178
f1_keywords:
- vbaxl10.chm144178
ms.prod: excel
api_name:
- Excel.Range.Precedents
ms.assetid: 3c00cfb4-1c12-668d-a952-89f9b1ef129f
ms.date: 06/08/2017
---


# Range.Precedents Property (Excel)

Returns a  **[Range](range-object-excel.md)** object that represents all the precedents of a cell. This can be a multiple selection (a union of **Range** objects) if there's more than one precedent. Read-only.


## Syntax

 _expression_ . **Precedents**

 _expression_ A variable that represents a **Range** object.


## Example

This example selects the precedents of cell A1 on Sheet1.


```vb
Worksheets("Sheet1").Activate 
Range("A1").Precedents.Select
```


## See also


#### Concepts


[Range Object](range-object-excel.md)

