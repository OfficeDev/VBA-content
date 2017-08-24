---
title: ShapeRange.Flip Method (Publisher)
keywords: vbapb10.chm2293781
f1_keywords:
- vbapb10.chm2293781
ms.prod: publisher
api_name:
- Publisher.ShapeRange.Flip
ms.assetid: fad24b08-9ada-0d6f-f526-ceec9ef996c1
ms.date: 06/08/2017
---


# ShapeRange.Flip Method (Publisher)

Flips the specified shape around its horizontal or vertical axis, or flips all the shapes in the specified shape range around their horizontal or vertical axes.


## Syntax

 _expression_. **Flip**( **_FlipCmd_**)

 _expression_A variable that represents a  **ShapeRange** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|FlipCmd|Required| **MsoFlipCmd**| Specifies whether the shape is flipped horizontally or vertically.|

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


