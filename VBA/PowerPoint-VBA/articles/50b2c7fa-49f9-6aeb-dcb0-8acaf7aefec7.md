
# TextFrame2.HasText Property (PowerPoint)

 Indicates whether the shape that contains the specified text frame has text associated with it. Read-only.


## Syntax

 _expression_. **HasText**

 _expression_ An expression that returns a **TextFrame2** object.


### Return Value

MsoTriState


## Remarks

The value of the  **HasText** property can be one of the following **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|The specified text frame does not have text.|
|**msoTrue**| The specified text frame has text.|

## Example

The followin example tests whether shape two on slide one contains text, and if it does, resizes the shape to fit the text.


```vb
Public Sub HasText_Example()



    Dim pptSlide As Slide

    Set pptSlide = ActivePresentation.Slides(1)

    With pptSlide.Shapes(2).TextFrame

        If .HasText Then .AutoSize = ppAutoSizeShapeToFitText

    End With



End Sub
```


## See also


#### Concepts


[TextFrame2 Object](ae017598-8330-4673-db1a-53b284acb709.md)
#### Other resources


[TextFrame2 Object Members](bce672a4-b108-b223-7e65-71f07d7f4197.md)
