---
title: CalculatedItems.Item Method (Excel)
keywords: vbaxl10.chm250075
f1_keywords:
- vbaxl10.chm250075
ms.prod: excel
api_name:
- Excel.CalculatedItems.Item
ms.assetid: ad7642b5-2579-17b4-ed2f-ebcac54bb595
ms.date: 06/08/2017
---


# CalculatedItems.Item Method (Excel)

Returns a single object from a collection.


## Syntax

 _expression_ . **Item**( **_Index_** )

 _expression_ A variable that represents a **CalculatedItems** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|The name or index number for the object.|

### Return Value

A  **PivotItem** object contained by the collection.


## Remarks

The text name of the object is the value of the  **[Name](pivotitem-name-property-excel.md)** and **[Value](pivotitem-value-property-excel.md)** properties.


## Example

This example hides calculated item one.


```vb
Worksheets(1).PivotTables(1).PivotFields("year") _ 
 .CalculatedItems.Item(1).Visible = False
```


## See also


#### Concepts


[CalculatedItems Collection](calculateditems-object-excel.md)

