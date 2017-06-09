---
title: CalculatedFields.Item Method (Excel)
keywords: vbaxl10.chm244075
f1_keywords:
- vbaxl10.chm244075
ms.prod: excel
api_name:
- Excel.CalculatedFields.Item
ms.assetid: cae0c3a5-3403-f1b1-fe7f-c38ff6be6b07
ms.date: 06/08/2017
---


# CalculatedFields.Item Method (Excel)

Returns a single object from a collection.


## Syntax

 _expression_ . **Item**( **_Index_** )

 _expression_ A variable that represents a **CalculatedFields** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|The name or index number for the object.|

### Return Value

A  **PivotField** object contained by the collection.


## Remarks

The text name of the object is the value of the  **[Name](pivotfield-name-property-excel.md)** and **[Value](pivotfield-value-property-excel.md)** properties.


## Example

This example sets the formula for calculated field one.


```vb
Worksheets(1).PivotTables(1).CalculatedFields.Item(1) _ 
 .Formula = "=Revenue - Cost"
```


## See also


#### Concepts


[CalculatedFields Collection](calculatedfields-object-excel.md)

