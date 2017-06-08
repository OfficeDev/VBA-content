---
title: PropertyEffect.Property Property (PowerPoint)
keywords: vbapp10.chm662003
f1_keywords:
- vbapp10.chm662003
ms.prod: powerpoint
api_name:
- PowerPoint.PropertyEffect.Property
ms.assetid: bb0ef094-0edd-3bc4-c02a-70fc8646017e
ms.date: 06/08/2017
---


# PropertyEffect.Property Property (PowerPoint)

Sets or returns an  **[MsoAnimProperty](msoanimproperty-enumeration-powerpoint.md)** constant that represents an animation property. Read/write.


## Syntax

 _expression_. **Property**

 _expression_ A variable that represents a **PropertyEffect** object.


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


[PropertyEffect Object](propertyeffect-object-powerpoint.md)

