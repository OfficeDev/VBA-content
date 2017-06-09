---
title: ChartObjects.Item Method (Excel)
keywords: vbaxl10.chm497106
f1_keywords:
- vbaxl10.chm497106
ms.prod: excel
api_name:
- Excel.ChartObjects.Item
ms.assetid: 0dbc6680-73ee-73a8-c3d8-f05faf6dd596
ms.date: 06/08/2017
---


# ChartObjects.Item Method (Excel)

Returns a single object from a collection.


## Syntax

 _expression_ . **Item**( **_Index_** )

 _expression_ A variable that represents a **ChartObjects** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|The name or index number for the object.|

### Return Value

An Object value that represents an object contained by the collection.


## Remarks

The text name of the object is the value of the  **Name** and **Value** properties.


## Example

This example activates embedded chart one.


```vb
Worksheets("sheet1").ChartObjects.Item(1).Activate
```


## See also


#### Concepts


[ChartObjects Object](chartobjects-object-excel.md)

