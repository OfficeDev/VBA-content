---
title: ShapeRange.Item Method (Excel)
keywords: vbaxl10.chm640074
f1_keywords:
- vbaxl10.chm640074
ms.prod: excel
api_name:
- Excel.ShapeRange.Item
ms.assetid: a8458e74-5279-3e47-308f-6c0647c00ee9
ms.date: 06/08/2017
---


# ShapeRange.Item Method (Excel)

Returns a single object from a collection.


## Syntax

 _expression_ . **Item**( **_Index_** )

 _expression_ A variable that represents a **ShapeRange** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|The name or index number for the object.|

### Return Value

A  **[Shape](shape-object-excel.md)** object contained by the collection.


## Example

This example sets the  **OnAction** property for shape two in a shape range. If the sr variable doesn?t represent a **ShapeRange** object, this example fails.


```vb
Dim sr As Shape 
sr.Item(2).OnAction = "ShapeAction"
```


## See also


#### Concepts


[ShapeRange Object](shaperange-object-excel.md)

