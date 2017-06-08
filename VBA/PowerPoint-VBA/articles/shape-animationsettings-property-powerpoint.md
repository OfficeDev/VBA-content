---
title: Shape.AnimationSettings Property (PowerPoint)
keywords: vbapp10.chm547047
f1_keywords:
- vbapp10.chm547047
ms.prod: powerpoint
api_name:
- PowerPoint.Shape.AnimationSettings
ms.assetid: c960d0de-afb3-55f2-b6fb-e67779cc42d2
ms.date: 06/08/2017
---


# Shape.AnimationSettings Property (PowerPoint)

Returns an  **[AnimationSettings](animationsettings-object-powerpoint.md)** object that represents all the special effects you can apply to the animation of the specified shape. Read-only.


## Syntax

 _expression_. **AnimationSettings**

 _expression_ A variable that represents a **Shape** object.


### Return Value

AnimationSettings


## Example

This example sets shape one on slide two in the active presentation to fly in from the left when the slide is built.


```vb
With ActivePresentation.Slides(2).Shapes(1).AnimationSettings

    .EntryEffect = ppEffectFlyFromLeft

    .TextLevelEffect = ppAnimateByAllLevels

End With
```


## See also


#### Concepts


[Shape Object](shape-object-powerpoint.md)

