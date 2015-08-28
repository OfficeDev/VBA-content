
# SlideRange.Duplicate Method (PowerPoint)

 **Last modified:** July 28, 2015

Creates a duplicate of the specified  **SlideRange** object, adds the new range of slides to the **Slides** collection immediately after the slide range specified originally, and then returns a **SlideRange** object that represents the duplicate slides.

## Syntax

 _expression_. **Duplicate**

 _expression_A variable that represents a  **SlideRange** object.


### Return Value

SlideRange


## Example

This example creates a duplicate of slide one in the active presentation and then sets the background shading and the title text of the new slide. The new slide will be slide two in the presentation.


```
Set newSlide = ActivePresentation.Slides(1).Duplicate

With newSlide

    .Background.Fill.PresetGradient msoGradientVertical, _

        1, msoGradientGold

    .Shapes.Title.TextFrame.TextRange _

        .Text = "Second Quarter Earnings"

End With
```


## See also


#### Concepts


 [SlideRange Object](440ab59d-744a-209f-bf28-d0acd3a21e1a.md)
#### Other resources


 [SlideRange Object Members](f819c56d-96d5-836d-0d1f-49e505696f34.md)
