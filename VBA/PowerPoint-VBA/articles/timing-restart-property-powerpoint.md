---
title: Timing.Restart Property (PowerPoint)
keywords: vbapp10.chm653016
f1_keywords:
- vbapp10.chm653016
ms.prod: powerpoint
api_name:
- PowerPoint.Timing.Restart
ms.assetid: 0dd82d15-aa92-1de9-6406-957710c26fb6
ms.date: 06/08/2017
---


# Timing.Restart Property (PowerPoint)

Represents whether the animation effect restarts after the effect has started once. Read/write.


## Syntax

 _expression_. **Restart**

 _expression_ A variable that represents a **Timing** object.


### Return Value

MsoAnimEffectRestart


## Remarks

The value of the  **Restart** property can be one of these **MsoAnimEffectRestart** constants. The default is **msoAnimEffectRestartNever**.


||
|:-----|
|**msoAnimEffectRestartAlways**|
|**msoAnimEffectRestartNever**|
|**msoAnimEffectRestartWhenOff**|

## Example

The following example adds a shape and an animation to it, then sets the animation's restart behavior.


```vb
Sub AddShapeSetTiming()

    Dim effDiamond As Effect
    Dim shpRectangle As Shape

    'Adds shape and sets animation

    Set shpRectangle = ActivePresentation.Slides(1).Shapes _
        .AddShape(Type:=msoShapeRectangle, Left:=100, Top:=100, _
        Width:=50, Height:=50)

    Set effDiamond = ActivePresentation.Slides(1).TimeLine.MainSequence _
        .AddEffect(Shape:=shpRectangle, effectId:=msoAnimEffectPathDiamond)

    With effDiamond.Timing
        .Duration = 3
        .RepeatDuration = 5
        .RepeatCount = 3
        .Restart = msoAnimEffectRestartAlways
    End With

End Sub
```


## See also


#### Concepts


[Timing Object](timing-object-powerpoint.md)

