---
title: AnimationBehaviors Object (PowerPoint)
keywords: vbapp10.chm656000
f1_keywords:
- vbapp10.chm656000
ms.prod: powerpoint
api_name:
- PowerPoint.AnimationBehaviors
ms.assetid: 40e11093-5cbd-c8d3-04b5-4cd7de97bfa7
ms.date: 06/08/2017
---


# AnimationBehaviors Object (PowerPoint)

Represents a collection of  **[AnimationBehavior](animationbehavior-object-powerpoint.md)** objects.


## Example

Use the [Add](animationbehaviors-add-method-powerpoint.md)method to add an animation behavior. The following example adds a five-second animated rotation behavior to the main animation sequence on the first slide.


```vb
Sub AnimationObject()

    Dim timeMain As TimeLine



    'Reference the main animation timeline

    Set timeMain = ActivePresentation.Slides(1).TimeLine



    'Add a five-second animated rotation behavior

    'as the first animation in the main animation sequence

    timeMain.MainSequence(1).Behaviors.Add Type:=msoAnimTypeRotation, Index:=1

End Sub
```


## See also


#### Concepts


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)

