---
title: Shape.Flip Method (PowerPoint)
keywords: vbapp10.chm547004
f1_keywords:
- vbapp10.chm547004
ms.prod: powerpoint
api_name:
- PowerPoint.Shape.Flip
ms.assetid: f340183a-4ef6-1a17-bbbb-5b1ec2b9aa4e
ms.date: 06/08/2017
---


# Shape.Flip Method (PowerPoint)

Flips the specified shape around its horizontal or vertical axis.


## Syntax

 _expression_. **Flip**( **_FlipCmd_** )

 _expression_ A variable that represents a **Shape** object.


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


[Shape Object](shape-object-powerpoint.md)

