
# RulerLevels.Parent Property (PowerPoint)

 **Last modified:** July 28, 2015

Returns the parent object for the specified object.

## Syntax

 _expression_. **Parent**

 _expression_A variable that represents a  **RulerLevels** object.


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


 [RulerLevels Object](890f4bee-c48a-be48-2cac-b73736a5bdf0.md)
#### Other resources


 [RulerLevels Object Members](3c0f8fde-0956-eff6-0c3a-9c398f15f40a.md)
