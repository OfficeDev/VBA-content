---
title: ColorEffect.From Property (PowerPoint)
keywords: vbapp10.chm659004
f1_keywords:
- vbapp10.chm659004
ms.prod: powerpoint
api_name:
- PowerPoint.ColorEffect.From
ms.assetid: 177d8282-e374-3601-f0ab-63c9e48f5415
ms.date: 06/08/2017
---


# ColorEffect.From Property (PowerPoint)

Sets or returns a  **ColorFormat** object that represents the starting RGB color value of an animation behavior.


## Syntax

 _expression_. **From**

 _expression_ A variable that represents a **ColorEffect** object.


## Remarks

Use this property in conjunction with the  **[To](coloreffect-to-property-powerpoint.md)** property to transition from one color to another.

Do not confuse this property with the  **FromX** or **FromY** properties of the **[ScaleEffect](scaleeffect-object-powerpoint.md)** and **[MotionEffect](motioneffect-object-powerpoint.md)** objects, which are only used for scaling or motion effects.


## Example

The following example adds a color effect and immediately changes its color.


```vb
Sub AddAndChangeColorEffect() 
    Dim effBlinds As Effect 
    Dim tlnTiming As TimeLine 
    Dim shpRectangle As Shape 
    Dim animColorEffect As AnimationBehavior 
    Dim clrEffect As ColorEffect 
 
    'Adds rectangle and sets effect and animation 
    Set shpRectangle = ActivePresentation.Slides(1).Shapes _ 
        .AddShape(Type:=msoShapeRectangle, Left:=100, _ 
        Top:=100, Width:=50, Height:=50) 
    Set effBlinds = t.MainSequence.AddEffect(Shape:=shpRectangle, _ 
        effectId:=msoAnimEffectBlinds) 
    Set animColorEffect = tlnTimming.MainSequence(1).Behaviors _ 
        .Add(Type:=msoAnimTypeColor) 
    Set clrEffect = animColorEffect.ColorEffect 
 
    'Sets the animation effect starting and ending colors 
    clrEffect.From.RGB = RGB(Red:=255, Green:=255, Blue:=0) 
    clrEffect.To.RGB = RGB(Red:=0, Green:=255, Blue:=255) 
End Sub
```


## See also


#### Concepts


[ColorEffect Object](coloreffect-object-powerpoint.md)

