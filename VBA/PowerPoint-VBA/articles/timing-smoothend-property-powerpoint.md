---
title: Timing.SmoothEnd Property (PowerPoint)
keywords: vbapp10.chm653014
f1_keywords:
- vbapp10.chm653014
ms.prod: powerpoint
api_name:
- PowerPoint.Timing.SmoothEnd
ms.assetid: 4d5d746b-ed5f-e708-287f-62e02684040c
ms.date: 06/08/2017
---


# Timing.SmoothEnd Property (PowerPoint)

Determines whether an animation should decelerate as it ends. Read/write.


## Syntax

 _expression_. **SmoothEnd**

 _expression_ A variable that represents a **Timing** object.


### Return Value

MsoTriState


## Remarks

The value of the  **SmoothEnd** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**| The default. An animation does not decelerate when it ends.|
|**msoTrue**| An animation decelerates when it ends.|

## Example

The following example adds a shape to a slide, animates the shape, and instructs the shape to decelerate when it ends.


```vb
Sub AddShapeSetTiming()

    Dim effDiamond As Effect
    Dim shpRectangle As Shape

    'Adds shape and sets animation effect
    Set shpRectangle = ActivePresentation.Slides(1).Shapes _
        .AddShape(Type:=msoShapeRectangle, Left:=100, _
        Top:=100, Width:=50, Height:=50)

    Set effDiamond = ActivePresentation.Slides(1).TimeLine.MainSequence _
        .AddEffect(Shape:=shpRectangle, effectId:=msoAnimEffectPathDiamond)

    'Sets duration of effect and slows animation at end
    With effDiamond.Timing
        .Duration = 5
        .SmoothEnd = msoTrue
    End With

End Sub
```


## See also


#### Concepts


[Timing Object](timing-object-powerpoint.md)

