
# EffectInformation.AnimateTextInReverse Property (PowerPoint)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Determines whether the specified shape is built in reverse order. Applies only to shapes (such as shapes containing lists) that can be built in more than one step. Read/write.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **AnimateTextInReverse**

 _expression_A variable that represents an  **EffectInformation** object.


### Return Value

MsoTriState


## Remarks
<a name="sectionSection1"> </a>

The value of the  **AnimateTextInReverse Property** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
| **msoFalse**| The specified shape is not built in reverse order.|
| **msoTrue**| The specified shape is built in reverse order.|
You do not see the effects of setting this property unless the specified shape gets animated. For a shape to be animated, the  **TextLevelEffect** property of the **AnimationSettings** object for the shape must be set to something other than **ppAnimateLevelNone** and the ** [Animate](7434630f-3c73-4261-36f7-a26d45e9df11.md)**property must be set to  **True**.


## Example
<a name="sectionSection2"> </a>

This example adds a slide after slide one in the active presentation, sets the title text, adds a three-item list to the text placeholder, and sets the list to be built in reverse order.


```
With ActivePresentation.Slides.Add(2, ppLayoutText).Shapes

    .Item(1).TextFrame.TextRange.Text = "Top Three Reasons"

    With .Item(2)

        .TextFrame.TextRange = "Reason 1" &amp; Chr(13) _

            &amp; "Reason 2" &amp; Chr(13) &amp; "Reason 3"

        With .AnimationSettings

            .Animate = msoTrue

            .TextLevelEffect = ppAnimateByFirstLevel

            .AnimateTextInReverse = msoTrue

        End With

    End With

End With
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [EffectInformation Object](9b3d09f4-229b-8392-f9a4-777bf6557632.md)
 [EffectInformation Object Members](a4d1a670-2592-5b92-9506-2e576b3a4e88.md)
