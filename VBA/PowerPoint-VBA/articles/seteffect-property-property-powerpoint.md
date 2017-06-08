---
title: SetEffect.Property Property (PowerPoint)
keywords: vbapp10.chm670003
f1_keywords:
- vbapp10.chm670003
ms.prod: powerpoint
api_name:
- PowerPoint.SetEffect.Property
ms.assetid: 75f31c60-327d-ce11-2703-d05ed870ef1b
ms.date: 06/08/2017
---


# SetEffect.Property Property (PowerPoint)

Sets or returns an  **[MsoAnimProperty](msoanimproperty-enumeration-powerpoint.md)** constant that represents an animation property. Read/write.


## Syntax

 _expression_. **Property**

 _expression_ A variable that represents a **SetEffect** object.


### Return Value

MsoAnimProperty


## Example

The following example adds a shape, adds a three-second fill animation to that shape, and sets the fill animation to color.


```vb
Sub AddShapeSetAnimFill()

    Dim effBlinds As Effect
    Dim shpRectangle As Shape
    Dim animProperty As AnimationBehavior

    Set shpRectangle = ActivePresentation.Slides(1).Shapes _
        .AddShape(Type:=msoShapeRectangle, Left:=100, _
        Top:=100, Width:=50, Height:=50)

    Set effBlinds = ActivePresentation.Slides(1).TimeLine.MainSequence _
        .AddEffect(Shape:=shpRectangle, effectId:=msoAnimEffectBlinds)

    effBlinds.Timing.Duration = 3
    Set animProperty = effBlinds.Behaviors.Add(msoAnimTypeProperty)

    With animProperty.PropertyEffect
        .Property = msoAnimColor
        .From = RGB(Red:=0, Green:=0, Blue:=255)
        .To = RGB(Red:=255, Green:=0, Blue:=0)
    End With

End Sub
```


## See also


#### Concepts


[SetEffect Object](seteffect-object-powerpoint.md)

