---
title: PivotItemList.Item Method (Excel)
keywords: vbaxl10.chm721074
f1_keywords:
- vbaxl10.chm721074
ms.prod: excel
api_name:
- Excel.PivotItemList.Item
ms.assetid: 69d0c71b-aa5a-b6cd-41d7-825197af869e
ms.date: 06/08/2017
---


# PivotItemList.Item Method (Excel)

Returns a single object from a collection.


## Syntax

 _expression_ . **Item**( **_Index_** )

 _expression_ A variable that represents a **PivotItemList** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|The name or index number for the object.|

### Return Value

A  **[PivotItem](pivotitem-object-excel.md)** object contained by the collection.


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


[PivotItemList Object](pivotitemlist-object-excel.md)

