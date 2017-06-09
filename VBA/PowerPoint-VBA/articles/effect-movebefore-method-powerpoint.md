---
title: Effect.MoveBefore Method (PowerPoint)
keywords: vbapp10.chm652005
f1_keywords:
- vbapp10.chm652005
ms.prod: powerpoint
api_name:
- PowerPoint.Effect.MoveBefore
ms.assetid: c71f8785-737d-b2cf-8d9d-bed49e1ba754
ms.date: 06/08/2017
---


# Effect.MoveBefore Method (PowerPoint)

Moves one animation effect to before another animation effect.


## Syntax

 _expression_. **MoveBefore**( **_Effect_** )

 _expression_ A variable that represents an **Effect** object.


## Example

The following example moves one effect in front of another one.


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

    effTwo.MoveBefore Effect:=effOne

End Sub
```


## See also


#### Concepts


[Effect Object](effect-object-powerpoint.md)


