---
title: SlideShowTransition.Duration Property (PowerPoint)
keywords: vbapp10.chm539011
f1_keywords:
- vbapp10.chm539011
ms.prod: powerpoint
api_name:
- PowerPoint.SlideShowTransition.Duration
ms.assetid: f8c47dda-9687-e437-8038-dae11c022914
ms.date: 06/08/2017
---


# SlideShowTransition.Duration Property (PowerPoint)

Returns or sets the length of an animation in seconds. Read/write.


## Syntax

 _expression_. **Duration**

 _expression_ A variable that represents a **Timing** object.


### Return Value

Single


## Example

The following example adds a shape and an animation to that shape, then sets its animation duration.


```vb
Sub AddShapeSetTiming()

    Dim effDiamond As Effect
    Dim shpRectangle As Shape

    'Adds shape and sets animation effect
    Set shpRectangle = ActivePresentation.Slides(1).Shapes _
        .AddShape(Type:=msoShapeRectangle, Left:=100, _
        Top:=100, Width:=50, Height:=50)

    Set effDiamond = ActivePresentation.Slides(1).TimeLine.MainSequence _
       .AddEffect(Shape:=sh, effectId:=msoAnimEffectPathDiamond)

    'Sets duration of effect
    effDiamond.Timing.Duration = 5

End Sub
```


## See also


#### Concepts


[Timing Object](timing-object-powerpoint.md)
[SlideShowTransition Object](slideshowtransition-object-powerpoint.md)

