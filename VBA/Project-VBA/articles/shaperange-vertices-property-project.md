---
title: ShapeRange.Vertices Property (Project)
ms.prod: project-server
ms.assetid: 5df31583-7e8a-2bc1-ed6b-719960fb7de1
ms.date: 06/08/2017
---


# ShapeRange.Vertices Property (Project)
Gets the coordinates of the vertices (and control points for a B?zier curve) as a series of coordinate pairs, for a shape range that contains a drawing. Read-only  **Variant**.

## Syntax

 _expression_. **Vertices**

 _expression_ A variable that represents a **ShapeRange** object.


## Remarks

You can use the array returned by the  **Vertices** property as an argument for the[AddCurve](shapes-addcurve-method-project.md) method or the[AddPolyLine](shapes-addpolyline-method-project.md) method.

For an array of vertices named  `vertArray`, the following table shows how the  **Vertices** property associates values in the array with the coordinates of vertices in a triangle.



|**Element in the array**|**Value of the element (in points)**|
|:-----|:-----|
| `vertArray(1, 1)`|The horizontal distance from the first vertex to the left side of the document.|
| `vertArray(1, 2)`|The vertical distance from the first vertex to the top of the document.|
| `vertArray(2, 1)`|The horizontal distance from the second vertex to the left side of the document.|
| `vertArray(2, 2)`|The vertical distance from the second vertex to the top of the document.|
| `vertArray(3, 1)`|The horizontal distance from the third vertex to the left side of the document.|
| `vertArray(3, 2)`|The vertical distance from the third vertex to the top of the document.|

## Property value

 **VARIANT**


## See also


#### Other resources


[ShapeRange Object](shaperange-object-project.md)
