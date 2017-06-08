---
title: PivotFormulas.Item Method (Excel)
keywords: vbaxl10.chm233075
f1_keywords:
- vbaxl10.chm233075
ms.prod: excel
api_name:
- Excel.PivotFormulas.Item
ms.assetid: 023f5702-9e18-f5d1-82b8-2603a98eb0b2
ms.date: 06/08/2017
---


# PivotFormulas.Item Method (Excel)

Returns a single object from a collection.


## Syntax

 _expression_ . **Item**( **_Index_** )

 _expression_ A variable that represents a **PivotFormulas** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|The name or index number for the object.|

### Return Value

A  **[PivotFormula](pivotformula-object-excel.md)** object contained by the collection.


## Example

This example displays the first formula for PivotTable one on worksheet one.


```vb
MsgBox Worksheets(1).PivotTables(1).PivotFormulas.Item(1).Formula
```


## See also


#### Concepts


[PivotFormulas Object](pivotformulas-object-excel.md)

