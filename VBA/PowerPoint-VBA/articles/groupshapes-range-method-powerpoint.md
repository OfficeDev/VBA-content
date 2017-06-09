---
title: GroupShapes.Range Method (PowerPoint)
keywords: vbapp10.chm549005
f1_keywords:
- vbapp10.chm549005
ms.prod: powerpoint
api_name:
- PowerPoint.GroupShapes.Range
ms.assetid: d7273a15-71f2-2e50-a481-055e8cc39e1f
ms.date: 06/08/2017
---


# GroupShapes.Range Method (PowerPoint)

Returns a  **ShapeRange** object.


## Syntax

 _expression_. **Range**( **_Index_** )

 _expression_ A variable that represents a **GroupShapes** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required|**Variant**|The individual shapes that are to be included in the range. Can be an  **Integer** that specifies the index number of the shape, a **String** that specifies the name of the shape, or an array that contains either integers or strings. If this argument is omitted, the **Range** method returns all the objects in the specified collection.|

### Return Value

ShapeRange


## Remarks

Although you can use the  **Range** method to return any number of shapes or slides, it is simpler to use the **Item** method if you only want to return a single member of the collection. For example, `Shapes(1)` is simpler than `Shapes.Range(1)`, and  `Slides(2)` is simpler than `Slides.Range(2)`.

To specify an array of integers or strings for  **Index**, you can use the **Array** function. For example, the following instruction returns two shapes specified by name.

 `Dim myArray() As Variant, myRange As Object myArray = Array("Oval 4", "Rectangle 5") Set myRange = ActivePresentation.Slides(1).Shapes.Range(myArray)`


## See also


#### Concepts


[GroupShapes Object](groupshapes-object-powerpoint.md)

