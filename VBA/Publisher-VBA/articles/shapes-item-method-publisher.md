---
title: Shapes.Item Method (Publisher)
keywords: vbapb10.chm2162688
f1_keywords:
- vbapb10.chm2162688
ms.prod: publisher
api_name:
- Publisher.Shapes.Item
ms.assetid: 174bbabb-e19f-4638-6dd8-780a8617fd70
ms.date: 06/08/2017
---


# Shapes.Item Method (Publisher)

Returns an individual object in a specified collection.


## Syntax

 _expression_. **Item**( **_Index_**)

 _expression_A variable that represents a  **Shapes** object.


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


