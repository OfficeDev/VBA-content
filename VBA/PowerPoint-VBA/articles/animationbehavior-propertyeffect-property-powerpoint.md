---
title: AnimationBehavior.PropertyEffect Property (PowerPoint)
keywords: vbapp10.chm657010
f1_keywords:
- vbapp10.chm657010
ms.prod: powerpoint
api_name:
- PowerPoint.AnimationBehavior.PropertyEffect
ms.assetid: a053462c-6ff6-52b4-2852-def0528780b2
ms.date: 06/08/2017
---


# AnimationBehavior.PropertyEffect Property (PowerPoint)

Returns a  **[PropertyEffect](propertyeffect-object-powerpoint.md)** object for a given animation behavior. Read-only.


## Syntax

 _expression_. **PropertyEffect**

 _expression_ A variable that represents an **AnimationBehavior** object.


### Return Value

PropertyEffect


## Example

The following example adds a shape with an effect to the active presentation and sets the animation effect properties for the shape to change colors.


```vb
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
        .From = RGB(Red:=0, Green:=0, Blue:=255)
        .To = RGB(Red:=255, Green:=0, Blue:=0)
    End With

End Sub
```


## See also


#### Concepts


[AnimationBehavior Object](animationbehavior-object-powerpoint.md)

