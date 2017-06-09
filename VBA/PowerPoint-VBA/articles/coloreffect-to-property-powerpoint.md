---
title: ColorEffect.To Property (PowerPoint)
keywords: vbapp10.chm659005
f1_keywords:
- vbapp10.chm659005
ms.prod: powerpoint
api_name:
- PowerPoint.ColorEffect.To
ms.assetid: c5a3a2bd-c33a-13ed-b2fd-e9ebb1f446e1
ms.date: 06/08/2017
---


# ColorEffect.To Property (PowerPoint)

Sets or returns a  **ColorFormat** object that represents the RGB color value of an animation behavior. Read/write.


## Syntax

 _expression_. **To**

 _expression_ A variable that represents a **ColorEffect** object.


### Return Value

ColorFormat


## Remarks

Use this property in conjunction with the  **From** property to transition from one color to another.

Do not confuse this property with the  **ToX** or **ToY** properties of the **[ScaleEffect](scaleeffect-object-powerpoint.md)** and **[MotionEffect](motioneffect-object-powerpoint.md)** objects, which are only used for scaling or motion effects.


## Example

The following example adds a color effect and changes its color from a light bluish green to yellow.


```vb
Sub AddAndChangeColorEffect()

    Dim effBlinds As Effect
    Dim tmlTiming As TimeLine
    Dim shpRectangle As Shape
    Dim animColor As AnimationBehavior
    Dim clrEffect As ColorEffect

    Set shpRectangle = ActivePresentation.Slides(1).Shapes _
        .AddShape(Type:=msoShapeRectangle, Left:=100, _
        Top:=100, Width:=50, Height:=50)

    Set tmlTiming = ActivePresentation.Slides(1).TimeLine
    Set effBlinds = tmlTiming.MainSequence.AddEffect(Shape:=shpRectangle, _
        effectId:=msoAnimEffectBlinds)

    Set animColor = tmlTiming.MainSequence(1).Behaviors _
        .Add(Type:=msoAnimTypeColor)

    Set clrEffect = animColor.ColorEffect
    clrEffect.From.RGB = RGB(Red:=255, Green:=255, Blue:=0)
    clrEffect.To.RGB = RGB(Red:=0, Green:=255, Blue:=255)

End Sub
```


## See also


#### Concepts


[ColorEffect Object](coloreffect-object-powerpoint.md)

