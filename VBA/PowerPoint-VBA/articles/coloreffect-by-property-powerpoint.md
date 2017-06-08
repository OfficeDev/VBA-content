---
title: ColorEffect.By Property (PowerPoint)
keywords: vbapp10.chm659003
f1_keywords:
- vbapp10.chm659003
ms.prod: powerpoint
api_name:
- PowerPoint.ColorEffect.By
ms.assetid: f0b841f0-694b-7cf0-fe71-1e54d840c099
ms.date: 06/08/2017
---


# ColorEffect.By Property (PowerPoint)

Returns a  **ColorFormat** object that represents a change to the color of the object by the specified number, expressed in RGB format. Read-only.


## Syntax

 _expression_. **By**

 _expression_ A variable that represents a **ColorEffect** object.


## Remarks

Do not confuse this property with the  **ByX** or **ByY** properties of the **[ScaleEffect](scaleeffect-object-powerpoint.md)** and **[MotionEffect](motioneffect-object-powerpoint.md)** objects, which are only used for scaling or motion effects.


## Example

This example adds a color effect and changes its color. This example assumes there is at least one shape on the first slide of the active presentation.


```vb
Sub AddAndChangeColorEffect()

    Dim effBlinds As Effect
    Dim tmlnShape As TimeLine
    Dim shpShape As Shape
    Dim animBehavior As AnimationBehavior
    Dim clrEffect As ColorEffect

    'Sets shape, timing, and effect
    Set shpShape = ActivePresentation.Slides(1).Shapes(1)
    Set tmlnShape = ActivePresentation.Slides(1).TimeLine
    Set effBlinds = tmlnShape.MainSequence.AddEffect _
        (Shape:=shpShape, effectId:=msoAnimEffectBlinds)

    'Adds animation behavior and color effect
    Set animBehavior = tmlnShape.MainSequence(1).Behaviors _
        .Add(Type:=msoAnimTypeColor)
    Set clrEffect = animBehavior.ColorEffect

    'Specifies color
    clrEffect.By.RGB = RGB(Red:=255, Green:=0, Blue:=0)

End Sub
```


## See also


#### Concepts


[ColorEffect Object](coloreffect-object-powerpoint.md)

