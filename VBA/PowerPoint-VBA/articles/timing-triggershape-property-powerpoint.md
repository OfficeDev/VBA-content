---
title: Timing.TriggerShape Property (PowerPoint)
keywords: vbapp10.chm653006
f1_keywords:
- vbapp10.chm653006
ms.prod: powerpoint
api_name:
- PowerPoint.Timing.TriggerShape
ms.assetid: 0b9431d2-0cea-d279-4aa7-24dd145e987e
ms.date: 06/08/2017
---


# Timing.TriggerShape Property (PowerPoint)

Sets or returns a  **Shape** object that represents the shape associated with an animation trigger. Read/write.


## Syntax

 _expression_. **TriggerShape**

 _expression_ A variable that represents a **Timing** object.


### Return Value

Shape


## Example

The following example adds two shapes to a slide, adds an animation to a shape, and begins the animation when the other shape is clicked.


```vb
Sub AddShapeSetTiming()

    Dim effDiamond As Effect
    Dim shpRectangle As Shape

    Set shpOval = _
      ActivePresentation.Slides(1).Shapes. _
      AddShape(Type:=msoShapeOval, Left:=400, Top:=100, Width:=100, Height:=50)

    Set shpRectangle = ActivePresentation.Slides(1).Shapes.  _
      AddShape(Type:=msoShapeRectangle, Left:=100, Top:=100, Width:=50, Height:=50)

    Set effDiamond = ActivePresentation.Slides(1).TimeLine. _
      InteractiveSequences.Add().AddEffect(Shape:=shpRectangle,  _
      effectId:=msoAnimEffectPathDiamond, trigger:=msoAnimTriggerOnShapeClick)

    With effDiamond.Timing
        .Duration = 5
        .TriggerShape = shpOval
    End With

End Sub
```


## See also


#### Concepts


[Timing Object](timing-object-powerpoint.md)

