---
title: OLEObjects.Item Method (Excel)
keywords: vbaxl10.chm422090
f1_keywords:
- vbaxl10.chm422090
ms.prod: excel
api_name:
- Excel.OLEObjects.Item
ms.assetid: 781b29f3-dcac-2679-72c2-a8d5d6280661
ms.date: 06/08/2017
---


# OLEObjects.Item Method (Excel)

Returns a single object from a collection.


## Syntax

 _expression_ . **Item**( **_Index_** )

 _expression_ A variable that represents an **OLEObjects** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|The name or index number for the object.|

### Return Value

An Object value that represents an object contained by the collection.


## Remarks

The text name of the object is the value of the  **Name** and **Value** properties.


## Example

This example deletes OLE object one from Sheet1.


```vb
Worksheets("sheet1").OLEObjects.Item(1).Delete
```


## See also


#### Concepts


[OLEObjects Object](oleobjects-object-excel.md)

