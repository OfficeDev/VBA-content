---
title: Shapes.Item Method (PowerPoint)
keywords: vbapp10.chm543003
f1_keywords:
- vbapp10.chm543003
ms.prod: powerpoint
api_name:
- PowerPoint.Shapes.Item
ms.assetid: f6c5eac1-3b65-3023-3b7a-557c7bfb0f02
ms.date: 06/08/2017
---


# Shapes.Item Method (PowerPoint)

Returns a single  **Shape** object from the specified **Shapes** collection.


## Syntax

 _expression_. **Item**( **_Index_** )

 _expression_ A variable that represents a **Shapes** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required|**Variant**|The name or index number of the single  **Shape** object in the collection to be returned.|

### Return Value

Shape


## Example

This example sets the foreground color to red for the shape named "Rectangle 1" on slide one in the active presentation.


```vb
ActivePresentation.Slides.Item(1).Shapes.Item("rectangle 1").Fill _
    .ForeColor.RGB = RGB(128, 0, 0)
```


## See also


#### Concepts


[Shapes Object](shapes-object-powerpoint.md)

