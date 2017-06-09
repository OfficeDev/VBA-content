---
title: Shapes.AddCurve Method (PowerPoint)
keywords: vbapp10.chm543007
f1_keywords:
- vbapp10.chm543007
ms.prod: powerpoint
api_name:
- PowerPoint.Shapes.AddCurve
ms.assetid: 47f90182-a71b-a028-c43f-a85d59d2a56b
ms.date: 06/08/2017
---


# Shapes.AddCurve Method (PowerPoint)

Creates a B?zier curve. Returns a  **[Shape](shape-object-powerpoint.md)** object that represents the new curve.


## Syntax

 _expression_. **AddCurve**( **_SafeArrayOfPoints_** )

 _expression_ A variable that represents a **Shapes** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SafeArrayOfPoints_|Required|**Variant**|An array of coordinate pairs that specifies the vertices and control points of the curve. The first point you specify is the starting vertex, and the next two points are control points for the first B?zier segment. Then, for each additional segment of the curve, you specify a vertex and two control points. The last point you specify is the ending vertex for the curve. Note that you must always specify 3n + 1 points, where n is the number of segments in the curve.|

### Return Value

Shape


## Example

The following example adds a two-segment B?zier curve to myDocument.


```vb
Dim pts(1 To 7, 1 To 2) As Single

pts(1, 1) = 0

pts(1, 2) = 0

pts(2, 1) = 72

pts(2, 2) = 72

pts(3, 1) = 100

pts(3, 2) = 40

pts(4, 1) = 20

pts(4, 2) = 50

pts(5, 1) = 90

pts(5, 2) = 120

pts(6, 1) = 60

pts(6, 2) = 30

pts(7, 1) = 150

pts(7, 2) = 90

Set myDocument = ActivePresentation.Slides(1)

myDocument.Shapes.AddCurve SafeArrayOfPoints:=pts
```


## See also


#### Concepts


[Shapes Object](shapes-object-powerpoint.md)

