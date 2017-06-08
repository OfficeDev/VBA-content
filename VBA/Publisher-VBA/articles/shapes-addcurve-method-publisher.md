---
title: Shapes.AddCurve Method (Publisher)
keywords: vbapb10.chm2162706
f1_keywords:
- vbapb10.chm2162706
ms.prod: publisher
api_name:
- Publisher.Shapes.AddCurve
ms.assetid: 888a35cb-190d-4058-e0d7-a848d77ba920
ms.date: 06/08/2017
---


# Shapes.AddCurve Method (Publisher)

Adds a new  **[Shape](shape-object-publisher.md)** object representing a Bézier curve to the specified **[Shapes](shapes-object-publisher.md)** collection.


## Syntax

 _expression_. **AddCurve**( **_SafeArrayOfPoints_**)

 _expression_A variable that represents a  **Shapes** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|SafeArrayOfPoints|Required| **Variant**|An array of coordinate pairs that specifies the vertices and control points of the curve. The first point you specify is the starting vertex, and the next two points are control points for the first Bézier segment. Then, for each additional segment of the curve, you specify a vertex and two control points. The last point you specify is the ending vertex for the curve. Note that you must always specify 3n + 1 points, where n is the number of segments in the curve.|

### Return Value

Shape


## Remarks

For the array elements in  **_SafeArrayOfPoints_**, numeric values are evaluated in points; strings can be in any units supported by Microsoft Publisher (for example, "2.5 in").


## Example

The following example adds a two-segment Bézier curve to the first page of the active publication.


```vb
Dim shpCurve As Shape 
Dim arrPoints(1 To 4, 1 To 2) As Single 
 
arrPoints(1, 1) = 0 
arrPoints(1, 2) = 0 
arrPoints(2, 1) = 72 
arrPoints(2, 2) = 72 
arrPoints(3, 1) = 144 
arrPoints(3, 2) = 36 
arrPoints(4, 1) = 216 
arrPoints(4, 2) = 108 
 
Set shpCurve = ActiveDocument.Pages(1).Shapes.AddCurve _ 
 (SafeArrayOfPoints:=arrPoints)
```


