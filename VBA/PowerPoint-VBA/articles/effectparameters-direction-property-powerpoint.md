---
title: EffectParameters.Direction Property (PowerPoint)
keywords: vbapp10.chm654003
f1_keywords:
- vbapp10.chm654003
ms.prod: powerpoint
api_name:
- PowerPoint.EffectParameters.Direction
ms.assetid: 39ef6eaf-79fb-f70f-20fe-7ec07715d59c
ms.date: 06/08/2017
---


# EffectParameters.Direction Property (PowerPoint)

Determines the direction used for an animation effect. This property can be used only if the effect uses a direction. Read/write.


## Syntax

 _expression_. **Direction**

 _expression_ A variable that represents a **EffectParameters** object.


### Return Value

MsoAnimDirection


## Remarks

The value of the  **Direction** property can be one of these **MsoAnimDirection** constants.


||
|:-----|
|<strong>msoAnimDirectionAcross</strong>|
|
<strong>msoAnimDirectionBottom</strong>|
|
<strong>msoAnimDirectionBottomLeft</strong>|
|
<strong>msoAnimDirectionBottomRight</strong>|
|
<strong>msoAnimDirectionCenter</strong>|
|
<strong>msoAnimDirectionClockwise</strong>|
|
<strong>msoAnimDirectionCounterclockwise</strong>|
|
<strong>msoAnimDirectionCycleClockwise</strong>|
|
<strong>msoAnimDirectionCycleCounterclockwise</strong>|
|
<strong>msoAnimDirectionDown</strong>|
|
<strong>msoAnimDirectionDownLeft</strong>|
|
<strong>msoAnimDirectionDownRight</strong>|
|
<strong>msoAnimDirectionFontAllCaps</strong>|
|
<strong>msoAnimDirectionFontBold</strong>|
|
<strong>msoAnimDirectionFontItalic</strong>|
|
<strong>msoAnimDirectionFontShadow</strong>|
|
<strong>msoAnimDirectionFontStrikethrough</strong>|
|
<strong>msoAnimDirectionFontUnderline</strong>|
|
<strong>msoAnimDirectionGradual</strong>|
|
<strong>msoAnimDirectionHorizontal</strong>|
|
<strong>msoAnimDirectionHorizontalIn</strong>|
|
<strong>msoAnimDirectionHorizontalOut</strong>|
|
<strong>msoAnimDirectionIn</strong>|
|
<strong>msoAnimDirectionInBottom</strong>|
|
<strong>msoAnimDirectionInCenter</strong>|
|
<strong>msoAnimDirectionInSlightly</strong>|
|
<strong>msoAnimDirectionInstant</strong>|
|
<strong>msoAnimDirectionLeft</strong>|
|
<strong>msoAnimDirectionNone</strong>|
|
<strong>msoAnimDirectionOrdinalMask</strong>|
|
<strong>msoAnimDirectionOut</strong>|
|
<strong>msoAnimDirectionOutBottom</strong>|
|
<strong>msoAnimDirectionOutCenter</strong>|
|
<strong>msoAnimDirectionOutSlightly</strong>|
|
<strong>msoAnimDirectionRight</strong>|
|
<strong>msoAnimDirectionSlightly</strong>|
|
<strong>msoAnimDirectionTop</strong>|
|
<strong>msoAnimDirectionTopLeft</strong>|
|
<strong>msoAnimDirectionTopRight</strong>|
|
<strong>msoAnimDirectionUp</strong>|
|
<strong>msoAnimDirectionUpLeft</strong>|
|
<strong>msoAnimDirectionUpRight</strong>|
|
<strong>msoAnimDirectionVertical</strong>|
|
<strong>msoAnimDirectionVerticalIn</strong>|
|
<strong>msoAnimDirectionVerticalOut</strong>|

## Example

The following example adds a shape,and animates the shape to fly in from the left.


```vb
Sub AddShapeSetAnimFly()

    Dim effFly As Effect
    Dim shpCube As Shape

    Set shpCube = ActivePresentation.Slides(1).Shapes _
        .AddShape(Type:=msoShapeCube, Left:=100, _
        Top:=100, Width:=50, Height:=50)

    Set effFly = ActivePresentation.Slides(1).TimeLine.MainSequence _
        .AddEffect(Shape:=shpCube, effectId:=msoAnimEffectFly)

    effFly.Timing.Duration = 3
    effFly.EffectParameters.Direction = msoAnimDirectionLeft

End Sub
```


## See also


#### Concepts



[EffectParameters Object](effectparameters-object-powerpoint.md)

