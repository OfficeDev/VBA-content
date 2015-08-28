
# SlideShowWindows.Parent Property (PowerPoint)

 **Last modified:** July 28, 2015

Returns the parent object for the specified object.

## Syntax

 _expression_. **Parent**

 _expression_A variable that represents a  **SlideShowWindows** object.


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


 [SlideShowWindows Object](aa4c7a38-32ea-c206-ce1f-d78094410f52.md)
#### Other resources


 [SlideShowWindows Object Members](329da7bf-e6e5-fac2-1285-9758dcb45fe0.md)
