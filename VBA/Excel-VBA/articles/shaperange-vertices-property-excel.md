---
title: ShapeRange.Vertices Property (Excel)
keywords: vbaxl10.chm640120
f1_keywords:
- vbaxl10.chm640120
ms.prod: excel
api_name:
- Excel.ShapeRange.Vertices
ms.assetid: 16cee6dd-9403-70e7-14dc-e75860d6406d
ms.date: 06/08/2017
---


# ShapeRange.Vertices Property (Excel)

Returns the coordinates of the specified freeform drawing's vertices (and control points for B?zier curves) as a series of coordinate pairs. You can use the array returned by this property as an argument to the  **[AddCurve](shapes-addcurve-method-excel.md)** method or **[AddPolyLine](shapes-addpolyline-method-excel.md)** method. Read-only **Variant** .


## Syntax

 _expression_ . **Vertices**

 _expression_ A variable that represents a **ShapeRange** object.


## Remarks

The following table shows how the  **Vertices** property associates the values in the array `vertArray()` with the coordinates of a triangle's vertices.



|**vertArray element**|**Contains**|
|:-----|:-----|
| `vertArray(1, 1)`|The horizontal distance from the first vertex to the left side of the document|
| `vertArray(1, 2)`|The vertical distance from the first vertex to the top of the document|
| `vertArray(2, 1)`|The horizontal distance from the second vertex to the left side of the document|
| `vertArray(2, 2)`|The vertical distance from the second vertex to the top of the document|
| `vertArray(3, 1)`|The horizontal distance from the third vertex to the left side of the document|
| `vertArray(3, 2)`|The vertical distance from the third vertex to the top of the document|

## Example

This example assigns the vertex coordinates for shape one on  `myDocument` to the array variable `vertArray()` and displays the coordinates for the first vertex.


```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes(1) 
 vertArray = .Vertices 
 x1 = vertArray(1, 1) 
 y1 = vertArray(1, 2) 
 MsgBox "First vertex coordinates: " &; x1 &; ", " &; y1 
End With
```

This example creates a curve that has the same geometric description as shape one on  `myDocument`. Shape one must contain 3 _n_+1 vertices for this example to succeed.




```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes 
 .AddCurve .Item(1).Vertices 
End With
```


## See also


#### Concepts


[ShapeRange Object](shaperange-object-excel.md)

