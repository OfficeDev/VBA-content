
# ThreeDFormat.Depth Property (PowerPoint)

Returns or sets the depth of the shape's extrusion. Read/write.


## Syntax

 _expression_. **Depth**

 _expression_ A variable that represents a **ThreeDFormat** object.


### Return Value

Single


## Remarks

The  **Depth** property value can be from - 600 through 9600 (positive values produce an extrusion whose front face is the original shape; negative values produce an extrusion whose back face is the original shape).


## Example

This example adds an oval to  `myDocument`, and then specifies that the oval be extruded to a depth of 50 points and that the extrusion be purple.


```vb
Set myDocument = ActivePresentation.Slides(1)

Set myShape = myDocument.Shapes _
    .AddShape(msoShapeOval, 90, 90, 90, 40)

With myShape.ThreeD
    .Visible = True
    .Depth = 50
    'RGB value for purple
    .ExtrusionColor.RGB = RGB(255, 100, 255) 
End With
```


## See also


#### Concepts


[TickLabels Object](2ba878bf-3a76-1350-2bd4-615c2520f042.md)
[ThreeDFormat Object](d6eb7b36-57df-727e-fc5b-50b8c4790c1c.md)
#### Other resources


[TickLabels Object Members](6e05b351-b72c-9ef4-635a-f91c94781cb1.md)
[ThreeDFormat Object Members](8d24e2d8-6579-5a14-f403-aaa77b6ed0a6.md)
