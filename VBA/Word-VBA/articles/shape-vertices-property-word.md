---
title: Shape.Vertices Property (Word)
keywords: vbawd10.chm161480830
f1_keywords:
- vbawd10.chm161480830
ms.prod: word
api_name:
- Word.Shape.Vertices
ms.assetid: e51e17dd-9e4e-28ab-4efd-7913cab45ca9
ms.date: 06/08/2017
---


# Shape.Vertices Property (Word)

Returns the coordinates of the specified freeform drawing's vertices (and control points for B?zier curves) as a series of coordinate pairs. Read-only  **Variant** .


## Syntax

 _expression_ . **Vertices**

 _expression_ Required. A variable that represents a **[Shape](shape-object-word.md)** object.


## Remarks

You can use the array returned by this property as an argument for the  **AddCurve** or **AddPolyLine** method.

The following table shows how the  **Vertices** property associates values in the array _vertArray()_ with the coordinates of a triangle's vertices.



|**vertArray element**|**Contains**|
|:-----|:-----|
| `vertArray(1, 1)`|The horizontal distance from the first vertex to the left side of the document.|
| `vertArray(1, 2)`|The vertical distance from the first vertex to the top of the document.|
|vertArray(2, 1)|The horizontal distance from the second vertex to the left side of the document.|
| `vertArray(2, 2)`|The vertical distance from the second vertex to the top of the document.|
| `vertArray(3, 1)`|The horizontal distance from the third vertex to the left side of the document.|
| `vertArray(3, 2)`|The vertical distance from the third vertex to the top of the document.|

## Example

This example assigns the vertex coordinates for shape one in the active document to an array variable and displays the coordinates for the first vertex. Shape one must be a freeform drawing.


```vb
With ActiveDocument.Shapes(1) 
    vertArray = .Vertices 
    x1 = vertArray(1, 1) 
    y1 = vertArray(1, 2) 
    MsgBox "First vertex coordinates: " &; x1 &; ", " &; y1 
End With
```

This example creates a curve that has the same geometric description as shape one in the active document. This example assumes that the first shape is a B?zier curve containing 3n+1 vertices, where n is the number of curve segments.




```vb
With ActiveDocument.Shapes 
    .AddCurve .Item(1).Vertices, Selection.Range 
End With
```


## See also


#### Concepts


[Shape Object](shape-object-word.md)

