---
title: Effect.MoveTo Method (PowerPoint)
keywords: vbapp10.chm652004
f1_keywords:
- vbapp10.chm652004
ms.prod: powerpoint
api_name:
- PowerPoint.Effect.MoveTo
ms.assetid: 7b424225-e53c-7dc9-1e5c-14b824110027
ms.date: 06/08/2017
---


# Effect.MoveTo Method (PowerPoint)

Moves the specified object to a specific location within the same collection, renumbering all other items in the collection appropriately.


## Syntax

 _expression_. **MoveTo**( **_toPos_** )

 _expression_ A variable that represents an **Effect** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _toPos_|Required|**Long**|The index position to which to move the animation effect.|

## Example

This example moves an animation effect to the second in the animation effects collection for the specified shape.


```vb
Sub MoveEffect()

    Dim sldFirst as Slide
    Dim shpFirst As Shape
    Dim effAdd As Effect

    Set sldFirst = ActivePresentation.Slides(1)
    Set shpFirst = sldFirst.Shapes(1)
    Set effAdd = sldFirst.TimeLine.MainSequence.AddEffect _
        (Shape:=shpFirst, effectId:=msoAnimEffectBlinds)

    effAdd.MoveTo toPos:=2

End Sub
```

This example moves the second slide in the active presentation to the first slide.




```vb
Sub MoveSlideToNewLocation()
    ActivePresentation.Slides(2).MoveTo toPos:=1
End Sub
```


## See also


#### Concepts



[Effect Object](effect-object-powerpoint.md)

