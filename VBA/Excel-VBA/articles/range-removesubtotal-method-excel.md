---
title: Range.RemoveSubtotal Method (Excel)
keywords: vbaxl10.chm144185
f1_keywords:
- vbaxl10.chm144185
ms.prod: excel
api_name:
- Excel.Range.RemoveSubtotal
ms.assetid: ec1fd131-551d-009f-1eea-033d805bb34d
ms.date: 06/08/2017
---


# Range.RemoveSubtotal Method (Excel)

Removes subtotals from a list.


## Syntax

 _expression_ . **RemoveSubtotal**

 _expression_ A variable that represents a **Range** object.


### Return Value

Variant


## Example

This example removes subtotals from the range A1:G37 on Sheet1. The example should be run on a list that has subtotals.


```vb
Worksheets("Sheet1").Range("A1:G37").RemoveSubtotal
```


## See also


#### Concepts


[Range Object](range-object-excel.md)

