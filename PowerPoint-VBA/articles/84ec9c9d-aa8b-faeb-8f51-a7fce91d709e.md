
# AnimationPoint.Formula Property (PowerPoint)

 **Last modified:** July 28, 2015

Returns or sets a  **String** that represents a formula to use for calculating an animation. Read/write.

## Syntax

 _expression_. **Formula**

 _expression_A variable that represents a  **AnimationPoint** object.


### Return Value

String


## Example

The following example adds a shape, and adds a three-second fill animation to that shape.


```
Sub AddShapeSetAnimFill()



    Dim effBlinds As Effect

    Dim shpRectangle As Shape

    Dim animBlinds As AnimationBehavior



    'Adds rectangle and sets animiation effect

    Set shpRectangle = ActivePresentation.Slides(1).Shapes _

        .AddShape(Type:=msoShapeRectangle, Left:=100, _

        Top:=100, Width:=50, Height:=50)

    Set effBlinds = ActivePresentation.Slides(1).TimeLine.MainSequence _

        .AddEffect(Shape:=shpRectangle, effectId:=msoAnimEffectBlinds)



    'Sets the duration of the animation

    effBlinds.Timing.Duration = 3



    'Adds a behavior to the animation

    Set animBlinds = effBlinds.Behaviors.Add(msoAnimTypeProperty)



    'Sets the animation color effect and the formula to use

    With animBlinds.PropertyEffect

        .Property = msoAnimColor

        .Formula = RGB(Red:=255, Green:=255, Blue:=255)

    End With



End Sub
```


## See also


#### Concepts


 [AnimationPoint Object](79aa1a47-abab-f98f-955a-48be10a94c41.md)
#### Other resources


 [AnimationPoint Object Members](26acf251-6c44-f370-f7ac-48057352cec6.md)
