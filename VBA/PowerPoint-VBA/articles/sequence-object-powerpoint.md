---
title: Sequence Object (PowerPoint)
keywords: vbapp10.chm651000
f1_keywords:
- vbapp10.chm651000
ms.prod: powerpoint
api_name:
- PowerPoint.Sequence
ms.assetid: 37a5224f-2461-b575-acb6-6905bbb5136d
ms.date: 06/08/2017
---


# Sequence Object (PowerPoint)

Represents a collection of  **[Effect](effect-object-powerpoint.md)** objects for a slide's interactive animation sequences. The **Sequence** collection is a member of the **[Sequences](sequences-object-powerpoint.md)** collection.


## Example

Use the [MainSequence](timeline-mainsequence-property-powerpoint.md)property of the  **[TimeLine](timeline-object-powerpoint.md)** object to return a **Sequence** object.

Use the [AddEffect](sequence-addeffect-method-powerpoint.md)method to add a new  **Sequence** object. This example adds a shape and an animation sequence to the first shape on the first slide in the active presentation.




```vb
Sub NewEffect()

    Dim effNew As Effect
    Dim shpFirst As Shape

    Set shpFirst = ActivePresentation.Slides(1).Shapes(1)

    Set effNew = ActivePresentation.Slides(1).TimeLine.MainSequence.AddEffect _
        (Shape:=shpFirst, effectId:=msoAnimEffectBlinds)

End Sub
```


## See also


#### Concepts


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)

