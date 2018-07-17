---
title: Effect.MoveAfter Method (PowerPoint)
keywords: vbapp10.chm652006
f1_keywords:
- vbapp10.chm652006
ms.prod: powerpoint
api_name:
- PowerPoint.Effect.MoveAfter
ms.assetid: 1d19f90c-51a6-d9bd-5593-53c67c7df415
ms.date: 06/08/2017
---


# Effect.MoveAfter Method (PowerPoint)

Moves one animation effect to after another animation effect.


## Syntax

 _expression_. **MoveAfter**( **_Effect_** )

 _expression_ A variable that represents an **Effect** object.


## Example

The following example moves one effect to after another.


```vb
Sub MoveEffect()

    Dim effOne As Effect
    Dim effTwo As Effect
    Dim shpFirst As Shape

    Set shpFirst = ActivePresentation.Slides(1).Shapes(1)

    Set effOne = ActivePresentation.Slides(1).TimeLine.MainSequence.AddEffect _
        (Shape:=shpFirst, effectId:=msoAnimEffectBlinds)

    Set effTwo = ActivePresentation.Slides(1).TimeLine.MainSequence.AddEffect _
        (Shape:=shpFirst, effectId:=msoAnimEffectBlast)

    effOne.MoveAfter Effect:=effTwo

End Sub
```


## See also


#### Concepts



[Effect Object](effect-object-powerpoint.md)

