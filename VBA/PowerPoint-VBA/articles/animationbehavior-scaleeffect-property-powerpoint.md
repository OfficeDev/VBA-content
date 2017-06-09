---
title: AnimationBehavior.ScaleEffect Property (PowerPoint)
keywords: vbapp10.chm657008
f1_keywords:
- vbapp10.chm657008
ms.prod: powerpoint
api_name:
- PowerPoint.AnimationBehavior.ScaleEffect
ms.assetid: 8e8236ca-c389-a888-5e07-42101fb92126
ms.date: 06/08/2017
---


# AnimationBehavior.ScaleEffect Property (PowerPoint)

Returns a  **[ScaleEffect](scaleeffect-object-powerpoint.md)** object for a given animation behavior. Read-only.


## Syntax

 _expression_. **ScaleEffect**

 _expression_ A variable that represents an **AnimationBehavior** object.


### Return Value

ScaleEffect


## Example

This example scales the first shape on the first slide starting at zero and increasing in size until it reaches 100 percent of its original size.


```vb
Sub ChangeScale()

    Dim shpFirst As Shape
    Dim effNew As Effect
    Dim aniScale As AnimationBehavior

    Set shpFirst = ActivePresentation.Slides(1).Shapes(1)
    Set effNew = ActivePresentation.Slides(1).TimeLine.MainSequence _
        .AddEffect(Shape:=shpFirst, effectId:=msoAnimEffectCustom)

    Set aniScale = effNew.Behaviors.Add(msoAnimTypeScale)
    With aniScale.ScaleEffect
        'Starting size
        .FromX = 0
        .FromY = 0

        'Size after scale effect
        .ToX = 100
        .ToY = 100
    End With
	
End Sub
```


## See also


#### Concepts


[AnimationBehavior Object](animationbehavior-object-powerpoint.md)

