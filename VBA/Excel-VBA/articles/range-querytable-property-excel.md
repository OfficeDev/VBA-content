---
title: Range.QueryTable Property (Excel)
keywords: vbaxl10.chm144183
f1_keywords:
- vbaxl10.chm144183
ms.prod: excel
api_name:
- Excel.Range.QueryTable
ms.assetid: 6370d43c-74b5-1bb9-f849-c70006432504
ms.date: 06/08/2017
---


# Range.QueryTable Property (Excel)

Returns a  **[QueryTable](querytable-object-excel.md)** object that represents the query table that intersects the specified **[Range](range-object-excel.md)** object.


## Syntax

 _expression_ . **QueryTable**

 _expression_ A variable that represents a **Range** object.


## Example

This example refreshes the QueryTable object that intersects cell A10 on worksheet one.


```vb
Worksheets(1).Range("a10").QueryTable.Refresh
```


## See also


#### Concepts


[Range Object](range-object-excel.md)

