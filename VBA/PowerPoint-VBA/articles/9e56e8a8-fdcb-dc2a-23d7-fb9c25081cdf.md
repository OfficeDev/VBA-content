
# EffectInformation.AnimateTextInReverse Property (PowerPoint)

Determines whether the specified shape is built in reverse order. Applies only to shapes (such as shapes containing lists) that can be built in more than one step. Read/write.


## Syntax

 _expression_. **AnimateTextInReverse**

 _expression_ A variable that represents an **EffectInformation** object.


### Return Value

MsoTriState


## Remarks

The value of the  **AnimateTextInReverse Property** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**| The specified shape is not built in reverse order.|
|**msoTrue**| The specified shape is built in reverse order.|
You do not see the effects of setting this property unless the specified shape gets animated. For a shape to be animated, the  **TextLevelEffect** property of the **AnimationSettings** object for the shape must be set to something other than **ppAnimateLevelNone** and the **[Animate](7434630f-3c73-4261-36f7-a26d45e9df11.md)** property must be set to **True**.


## Example

This example adds a slide after slide one in the active presentation, sets the title text, adds a three-item list to the text placeholder, and sets the list to be built in reverse order.


```vb
With ActivePresentation.Slides.Add(2, ppLayoutText).Shapes
    .Item(1).TextFrame.TextRange.Text = "Top Three Reasons"
    With .Item(2)
        .TextFrame.TextRange = "Reason 1" &; Chr(13) _
            &; "Reason 2" &; Chr(13) &; "Reason 3"
        With .AnimationSettings
            .Animate = msoTrue
            .TextLevelEffect = ppAnimateByFirstLevel
            .AnimateTextInReverse = msoTrue
        End With
    End With
End With
```


## See also


#### Concepts


[EffectInformation Object](9b3d09f4-229b-8392-f9a4-777bf6557632.md)
[EffectInformation Object Members](a4d1a670-2592-5b92-9506-2e576b3a4e88.md)
