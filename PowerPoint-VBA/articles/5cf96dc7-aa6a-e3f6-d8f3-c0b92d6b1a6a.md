
# Font.Parent Property (PowerPoint)

 **Last modified:** July 28, 2015

Returns the parent object for the specified object.

## Syntax

 _expression_. **Parent**

 _expression_A variable that represents a  **Font** object.


### Return Value

Object


## Example

This example adds an oval containing text to slide one in the active presentation and rotates the oval and the text 45 degrees. The parent object for the text frame is the  **Shape** object that contains the text.


```
Set myShapes = ActivePresentation.Slides(1).Shapes

With myShapes.AddShape(Type:=msoShapeOval, Left:=50, _

        Top:=50, Width:=300, Height:=150).TextFrame

    .TextRange.Text = "Test text"

    .Parent.Rotation = 45

End With
```


## See also


#### Concepts


 [Font Object](ad62daaa-01a5-36cc-5451-e0da0134ac95.md)
#### Other resources


 [Font Object Members](a2043117-2222-dad3-d73c-0e9d5591c9be.md)
