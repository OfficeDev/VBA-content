---
title: PivotTables.Item Method (Excel)
keywords: vbaxl10.chm238074
f1_keywords:
- vbaxl10.chm238074
ms.prod: excel
api_name:
- Excel.PivotTables.Item
ms.assetid: 1bdc8558-ec67-2823-fd02-ecd5ae4ecee6
ms.date: 06/08/2017
---


# PivotTables.Item Method (Excel)

Returns a single object from a collection.


## Syntax

 _expression_ . **Item**( **_Index_** )

 _expression_ A variable that represents a **PivotTables** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|The name or index number for the object.|

### Return Value

A  **[PivotTable](pivottable-object-excel.md)** object contained by the collection.


## Remarks

The text name of the object is the value of the  **[Name](pivottable-name-property-excel.md)** and **[Value](pivottable-value-property-excel.md)** properties.


## Example

This example makes the Year field a row field in the first PivotTable report on Sheet3.


```vb
Worksheets("sheet3").PivotTables.Item(1) _ 
 .PivotFields("year").Orientation = xlRowField
```


## See also


#### Concepts


[PivotTables Object](pivottables-object-excel.md)

