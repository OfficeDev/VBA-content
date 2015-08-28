
# AnimationPoint.Time Property (PowerPoint)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Sets or returns the time at a given animation point. Read/write.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **Time**

 _expression_A variable that represents a  **SlideShowTransition** object.


### Return Value

Single


## Remarks
<a name="sectionSection1"> </a>

The value of the  **Time** property can be any floating-point value between 0 and 1, representing a percentage of the entire timeline from 0% to 100%. For example, a value of 0.2 would correspond to a point in time at 20% of the entire timeline duration from left to right.


## Example
<a name="sectionSection2"> </a>

This example inserts three fill color animation points in the main sequence animation timeline on the first slide.


```
Sub BuildTimeLine()



    Dim shpFirst As Shape

    Dim effMain As Effect

    Dim tmlMain As TimeLine

    Dim aniBhvr As AnimationBehavior

    Dim aniPoint As AnimationPoint



    Set shpFirst = ActivePresentation.Slides(1).Shapes(1)

    Set tmlMain = ActivePresentation.Slides(1).TimeLine

    Set effMain = tmlMain.MainSequence.AddEffect(Shape:=shpFirst, _

        EffectId:=msoAnimEffectBlinds)

    Set aniBhvr = tmlMain.MainSequence(1).Behaviors.Add _

        (Type:=msoAnimTypeProperty)



    With aniBhvr.PropertyEffect

        .Property = msoAnimShapeFillColor

        Set aniPoint = .Points.Add

        aniPoint.Time = 0.2

        aniPoint.Value = RGB(0, 0, 0)

        Set aniPoint = .Points.Add

        aniPoint.Time = 0.5

        aniPoint.Value = RGB(0, 255, 0)

        Set aniPoint = .Points.Add

        aniPoint.Time = 1

        aniPoint.Value = RGB(0, 255, 255)

    End With

End Sub
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [AnimationPoint Object](79aa1a47-abab-f98f-955a-48be10a94c41.md)
#### Other resources


 [AnimationPoint Object Members](26acf251-6c44-f370-f7ac-48057352cec6.md)
