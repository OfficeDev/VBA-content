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
|**msoAnimDirectionAcross**|
|**msoAnimDirectionBottom**|
|**msoAnimDirectionBottomLeft**|
|**msoAnimDirectionBottomRight**|
|**msoAnimDirectionCenter**|
|**msoAnimDirectionClockwise**|
|**msoAnimDirectionCounterclockwise**|
|**msoAnimDirectionCycleClockwise**|
|**msoAnimDirectionCycleCounterclockwise**|
|**msoAnimDirectionDown**|
|**msoAnimDirectionDownLeft**|
|**msoAnimDirectionDownRight**|
|**msoAnimDirectionFontAllCaps**|
|**msoAnimDirectionFontBold**|
|**msoAnimDirectionFontItalic**|
|**msoAnimDirectionFontShadow**|
|**msoAnimDirectionFontStrikethrough**|
|**msoAnimDirectionFontUnderline**|
|**msoAnimDirectionGradual**|
|**msoAnimDirectionHorizontal**|
|**msoAnimDirectionHorizontalIn**|
|**msoAnimDirectionHorizontalOut**|
|**msoAnimDirectionIn**|
|**msoAnimDirectionInBottom**|
|**msoAnimDirectionInCenter**|
|**msoAnimDirectionInSlightly**|
|**msoAnimDirectionInstant**|
|**msoAnimDirectionLeft**|
|**msoAnimDirectionNone**|
|**msoAnimDirectionOrdinalMask**|
|**msoAnimDirectionOut**|
|**msoAnimDirectionOutBottom**|
|**msoAnimDirectionOutCenter**|
|**msoAnimDirectionOutSlightly**|
|**msoAnimDirectionRight**|
|**msoAnimDirectionSlightly**|
|**msoAnimDirectionTop**|
|**msoAnimDirectionTopLeft**|
|**msoAnimDirectionTopRight**|
|**msoAnimDirectionUp**|
|**msoAnimDirectionUpLeft**|
|**msoAnimDirectionUpRight**|
|**msoAnimDirectionVertical**|
|**msoAnimDirectionVerticalIn**|
|**msoAnimDirectionVerticalOut**|

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

