
# FillFormat.Background Method (PowerPoint)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Specifies that the shape's fill should match the slide background. If you change the slide background after applying this method to a fill, the fill will also change.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **Background**

 _expression_A variable that represents a  **FillFormat** object.


## Remarks
<a name="sectionSection1"> </a>

Note that applying the  **Background** method to a shape's fill isn't the same as setting a transparent fill for the shape, nor is it always the same as applying the same fill to the shape as you apply to the background. The second example demonstrates this.


## Example
<a name="sectionSection2"> </a>

This example sets the fill of shape one on slide one in the active presentation to match the slide background.


```
ActivePresentation.Slides(1).Shapes(1).Fill.Background
```

This example sets the background for slide one in the active presentation to a preset gradient, adds a rectangle to the slide, and then places three ovals in front of the rectangle. The first oval has a fill that matches the slide background, the second has a transparent fill, and the third has the same fill applied to it as was applied to the background. Notice the difference in the appearances of these three ovals.




```
With ActivePresentation.Slides(1)

    .FollowMasterBackground = False

    .Background.Fill.PresetGradient _

        msoGradientHorizontal, 1, msoGradientDaybreak

    With .Shapes

        .AddShape msoShapeRectangle, 50, 200, 600, 100

        .AddShape(msoShapeOval, 75, 150, 150, 100) _

            .Fill.Background

        .AddShape(msoShapeOval, 275, 150, 150, 100).Fill _

            .Transparency = 1

        .AddShape(msoShapeOval, 475, 150, 150, 100) _

            .Fill.PresetGradient _

            msoGradientHorizontal, 1, msoGradientDaybreak

    End With

End With
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [FillFormat Object](5bd4e2cb-4466-b468-d494-bec30ed5c9d8.md)
#### Other resources


 [FillFormat Object Members](ccd26632-4ff8-6fad-2c5d-c26078eeff3b.md)
