
# Effect.Parent Property (PowerPoint)

 **Last modified:** July 28, 2015

Returns the parent object for the specified object.

## Syntax

 _expression_. **Parent**

 _expression_A variable that represents an  **Effect** object.


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


 [Effect Object Members](a110a644-1a87-b67c-b453-13c9d53004b7.md)
 [Effect Object](359ac3da-86cd-8003-d691-349d20fd1777.md)
