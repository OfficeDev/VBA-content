---
title: Shape.Flip Method (Publisher)
keywords: vbapb10.chm2228245
f1_keywords:
- vbapb10.chm2228245
ms.prod: publisher
api_name:
- Publisher.Shape.Flip
ms.assetid: 6d0004a5-2d76-955a-64ff-140dfbc313f3
ms.date: 06/08/2017
---


# Shape.Flip Method (Publisher)

Flips the specified shape around its horizontal or vertical axis, or flips all the shapes in the specified shape range around their horizontal or vertical axes.


## Syntax

 _expression_. **Flip**( **_FlipCmd_**)

 _expression_A variable that represents a  **Shape** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|FlipCmd|Required| **MsoFlipCmd**| Specifies whether the shape is flipped horizontally or vertically.|

### Return Value

Nothing


## Remarks

The FlipCmd parameter can be one of the following  **MsoFlipCmd** constants declared in the Microsoft Office type library.



| **msoFlipHorizontal**|
| **msoFlipVertical**|

## Example

This example adds a triangle to the first page of the active publication, duplicates the triangle, and then flips the duplicate triangle vertically and makes it red.


```vb
With ActiveDocument.Pages(1).Shapes _ 
 .AddShape(Type:=msoShapeRightTriangle, _ 
 Left:=10, Top:=10, Width:=50, Height:=50) _ 
 .Duplicate 
 .Fill.ForeColor.RGB = RGB(255, 0, 0) 
 .Flip msoFlipVertical 
End With 

```


