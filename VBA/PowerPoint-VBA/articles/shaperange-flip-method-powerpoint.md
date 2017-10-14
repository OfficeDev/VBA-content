---
title: ShapeRange.Flip Method (PowerPoint)
keywords: vbapp10.chm548004
f1_keywords:
- vbapp10.chm548004
ms.prod: powerpoint
api_name:
- PowerPoint.ShapeRange.Flip
ms.assetid: e9f5ceb5-2ddf-d70c-41d5-d5877043b62a
ms.date: 06/08/2017
---


# ShapeRange.Flip Method (PowerPoint)

Flips the specified shape range around its horizontal or vertical axis.


## Syntax

 _expression_. **Flip**( **_FlipCmd_** )

 _expression_ A variable that represents a **ShapeRange** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FlipCmd_|Required|**MsoFlipCmd**|Specifies whether the shape is to be flipped horizontally or vertically.|

## Remarks

The  _FlipCmd_ parameter value can be one of these **MsoFlipCmd** constants.


||
|:-----|
|**msoFlipHorizontal**|
|**msoFlipVertical**|

## Example

This example adds a triangle to  `myDocument`, duplicates the triangle, and then flips the duplicate triangle vertically and makes it red.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes _
        .AddShape(msoShapeRightTriangle, 10, 10, 50, 50).Duplicate
    .Fill.ForeColor.RGB = RGB(255, 0, 0)
    .Flip msoFlipVertical
End With
```


## See also


#### Concepts


[ShapeRange Object](shaperange-object-powerpoint.md)

