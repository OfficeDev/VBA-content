---
title: CanvasShapes.AddCurve Method (Word)
keywords: vbawd10.chm7536652
f1_keywords:
- vbawd10.chm7536652
ms.prod: word
api_name:
- Word.CanvasShapes.AddCurve
ms.assetid: d66512d3-77a7-4f52-ad94-44712c75e87d
ms.date: 06/08/2017
---


# CanvasShapes.AddCurve Method (Word)

Returns a  **[Shape](shape-object-word.md)** object that represents a B?zier curve in a drawing canvas.


## Syntax

 _expression_ . **AddCurve**( **_SafeArrayOfPoints_** )

 _expression_ Required. A variable that represents a **[CanvasShapes](canvasshapes-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SafeArrayOfPoints_|Required| **Variant**|An array of coordinate pairs that specifies the vertices and control points of the curve. The first point you specify is the starting vertex, and the next two points are control points for the first B?zier segment. Then, for each additional segment of the curve, you specify a vertex and two control points. The last point you specify is the ending vertex for the curve. Note that you must always specify 3n + 1 points, where n is the number of segments in the curve.|

## Example

This example adds a B?zier curve to a new drawing canvas.


```vb
Sub CanvasBezier() 
 
 Dim docNew As Document 
 Dim shpCanvas As Shape 
 Dim sngArray(1 To 7, 1 To 2) As Single 
 
 Set docNew = Documents.Add 
 
 'Create a new drawing canvas 
 Set shpCanvas = docNew.Shapes.AddCanvas(Left:=100, _ 
 Top:=100, Width:=300, Height:=50) 
 
 sngArray(1, 1) = 0 
 sngArray(1, 2) = 0 
 sngArray(2, 1) = 50 
 sngArray(2, 2) = 50 
 sngArray(3, 1) = 100 
 sngArray(3, 2) = 0 
 sngArray(4, 1) = 150 
 sngArray(4, 2) = 50 
 sngArray(5, 1) = 200 
 sngArray(5, 2) = 0 
 sngArray(6, 1) = 250 
 sngArray(6, 2) = 50 
 sngArray(7, 1) = 300 
 sngArray(7, 2) = 0 
 
 'Add Bezier curve to drawing canvas 
 shpCanvas.CanvasItems.AddCurve _ 
 SafeArrayOfPoints:=sngArray 
 
End Sub
```


## See also


#### Concepts


[CanvasShapes Collection](canvasshapes-object-word.md)

