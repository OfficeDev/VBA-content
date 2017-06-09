---
title: Sequence.FindFirstAnimationFor Method (PowerPoint)
keywords: vbapp10.chm651006
f1_keywords:
- vbapp10.chm651006
ms.prod: powerpoint
api_name:
- PowerPoint.Sequence.FindFirstAnimationFor
ms.assetid: 124dda8e-b93a-5d8a-06ba-30529cf5c6a0
ms.date: 06/08/2017
---


# Sequence.FindFirstAnimationFor Method (PowerPoint)

Returns an  **[Effect](effect-object-powerpoint.md)** object that represents the first animation for a given shape.


## Syntax

 _expression_. **FindFirstAnimationFor**( **_Shape_** )

 _expression_ A variable that represents a **Sequence** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Shape_|Required|**[Shape](shape-object-powerpoint.md)**|The shape for which to find the first animation.|

### Return Value

Effect


## Example

The following example finds and deletes the first animation for a the first shape on the first slide. This example assumes that at least one animation effect exists for the specified shape.


```vb
Sub FindFirstAnimation()

    Dim sldFirst As Slide
    Dim shpFirst As Shape
    Dim effFirst As Effect

    Set sldFirst = ActivePresentation.Slides(1)
    Set shpFirst = sldFirst.Shapes(1)

    Set effFirst = sldFirst.TimeLine.MainSequence _
        .FindFirstAnimationFor(Shape:=shpFirst)

    effFirst.Delete

End Sub
```


## See also


#### Concepts


[Sequence Object](sequence-object-powerpoint.md)

