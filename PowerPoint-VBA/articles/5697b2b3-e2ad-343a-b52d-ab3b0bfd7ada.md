
# TabStops.Parent Property (PowerPoint)

 **Last modified:** July 28, 2015

Returns the parent object for the specified object.

## Syntax

 _expression_. **Parent**

 _expression_A variable that represents a  **TabStops** object.


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


 [TabStops Object](e23b36de-6a4d-84e5-bec1-8c3e0fd80c13.md)
#### Other resources


 [TabStops Object Members](62f6b7f4-45f8-108c-6294-8f24d3b2058c.md)
