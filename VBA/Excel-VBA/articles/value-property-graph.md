---
title: Value Property (Graph)
keywords: vbagr10.chm65542
f1_keywords:
- vbagr10.chm65542
ms.prod: excel
ms.assetid: c88258bc-7088-71df-87e7-49239239de76
ms.date: 06/08/2017
---


# Value Property (Graph)

Returns the value of the specified cell. If the cell is empty, Value returns the value Empty (use the IsEmpty function to test for this case). If the Range object contains more than one cell, this property returns an array of values (use the IsArray function to test for this case). Read/write Variant.

 _expression_. **Value**( **_RangeValueDataType_**)

 _expression_ Required. An expression that returns one of the objects in the Applies To list.

 **RangeValueDataType** Optional **Variant**.

## Example

This example sets the value of cell A1 on the datasheet to 3.14159.


```
myChart.Application.DataSheet.Range("A1").Value = 3.14159
```


