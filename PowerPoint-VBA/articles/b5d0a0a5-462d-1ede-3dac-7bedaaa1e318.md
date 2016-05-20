
# Shape.TextEffect Property (PowerPoint)

Returns a  **[TextEffectFormat](62434479-237f-01c4-712c-08e48b391d48.md)** object that contains text-effect formatting properties for the specified shape. Read-only.


## Syntax

 _expression_. **TextEffect**

 _expression_ A variable that represents a **Shape** object.


### Return Value

TextEffectFormat


## Remarks

Applies to  **[Shape](1da93849-99e0-827e-ced3-c6cf7f8569f3.md)** objects that represent WordArt.


## Example

This example sets the font style to bold for shape three on  `myDocument` if the shape is WordArt.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes(3)

    If .Type = msoTextEffect Then

        .TextEffect.FontBold = True

    End If

End With
```


## See also


#### Concepts


[Shape Object](1da93849-99e0-827e-ced3-c6cf7f8569f3.md)
#### Other resources


[Shape Object Members](e371c375-c16a-33ef-32b7-6dcb99d3d128.md)
