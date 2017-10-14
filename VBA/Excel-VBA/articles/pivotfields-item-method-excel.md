---
title: PivotFields.Item Method (Excel)
keywords: vbaxl10.chm242075
f1_keywords:
- vbaxl10.chm242075
ms.prod: excel
api_name:
- Excel.PivotFields.Item
ms.assetid: 497c8536-30cb-8c7b-8d83-62ae94a37a7f
ms.date: 06/08/2017
---


# PivotFields.Item Method (Excel)

Returns a single object from a collection.


## Syntax

 _expression_ . **Item**( **_Index_** )

 _expression_ A variable that represents a **PivotFields** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|The name or index number for the object.|

### Return Value

An Object value that represents an object contained by the collection.


## Remarks

The text name of the object is the value of the  **Name** and **Value** properties.


## Example

This example makes the Year field a row field in the first PivotTable report on Sheet3.


```vb
Worksheets("sheet3").PivotTables(1) _ 
 .PivotFields.Item("year").Orientation = xlRowField 

```


## See also


#### Concepts


[PivotFields Object](pivotfields-object-excel.md)

