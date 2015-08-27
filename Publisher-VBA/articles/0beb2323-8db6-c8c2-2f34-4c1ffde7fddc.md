
# ShapeRange.Vertices Property (Publisher)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Returns the coordinates of the specified freeform drawing's vertices (and control points for BÃ©zier curves) as a series of coordinate pairs. Read-only  **Variant**.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **Vertices**

 _expression_A variable that represents a  **ShapeRange** object.


## Remarks
<a name="sectionSection1"> </a>

You can use the array returned by this property as an argument to the  [AddCurve](888a35cb-190d-4058-e0d7-a848d77ba920.md)or  [AddPolyline](d49fb2bc-4df5-fff8-c741-2c0d35413fc5.md)methods.

The following table shows how the  **Vertices** property associates the values in the array `vertArray()` with the coordinates of a triangle's vertices.



|**vertArray element**|**Contains**|
|:-----|:-----|
| `vertArray(1, 1)`|The horizontal distance from the first vertex to the left side of the page.|
| `vertArray(1, 2)`|The vertical distance from the first vertex to the top of the page.|
| `vertArray(2, 1)`|The horizontal distance from the second vertex to the left side of the page.|
| `vertArray(2, 2)`|The vertical distance from the second vertex to the top of the page.|
| `vertArray(3, 1)`|The horizontal distance from the third vertex to the left side of the page.|
| `vertArray(3, 2)`|The vertical distance from the third vertex to the top of the page.|

## Example
<a name="sectionSection2"> </a>

This example assigns the vertex coordinates for shape one in the active publication to the array variable  `vertArray()` and displays the coordinates for the first vertex.


```
Dim vertArray As Variant 
Dim sngX1 As Single 
Dim sngY1 As Single 
 
With ActiveDocument.Pages(1).Shapes(1) 
 vertArray = .Vertices 
 sngX1 = vertArray(1, 1) 
 sngY1 = vertArray(1, 2) 
 MsgBox "First vertex coordinates: " &amp; sngX1 &amp; ", " &amp; sngY1 
End With
```

This example creates a curve that has the same geometric description as shape one in the active publication. Shape one must contain 3n+1 vertices for this example to work, where n is an integer greater than or equal to 1.




```
With ActiveDocument.Pages(1).Shapes 
 .AddCurve SafeArrayOfPoints:=.Item(1).Vertices 
End With 

```

