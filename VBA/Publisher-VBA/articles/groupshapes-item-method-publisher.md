---
title: GroupShapes.Item Method (Publisher)
keywords: vbapb10.chm3342336
f1_keywords:
- vbapb10.chm3342336
ms.prod: publisher
api_name:
- Publisher.GroupShapes.Item
ms.assetid: d0e2f8a6-6529-a274-410b-744c2bb55774
ms.date: 06/08/2017
---


# GroupShapes.Item Method (Publisher)

Returns an individual object in a specified collection.


## Syntax

 _expression_. **Item**( **_Index_**)

 _expression_A variable that represents a  **GroupShapes** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Index|Required| **Variant**|The number or name of the field or list box item to return.|

### Return Value

Shape


## Example

This example returns the first shape inside a grouped shape.


```vb
Dim shpTemp As Shape 
 
Set shpTemp = ActiveDocument.Pages(Index:=1) _ 
 .Shapes(1).GroupItems.Item(Index:=1)
```


