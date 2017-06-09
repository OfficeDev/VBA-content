---
title: Effect.Index Property (PowerPoint)
keywords: vbapp10.chm652008
f1_keywords:
- vbapp10.chm652008
ms.prod: powerpoint
api_name:
- PowerPoint.Effect.Index
ms.assetid: 1eac9295-e24c-c31e-3cd6-ace59f5ac04a
ms.date: 06/08/2017
---


# Effect.Index Property (PowerPoint)

Returns a  **Long** that represents the index number for an animation effect or design. Read-only.


## Syntax

 _expression_. **Index**

 _expression_ A variable that represents an **Effect** object.


### Return Value

Long


## Example

The following example displays the name and index number for all effects in the main animation sequence of the first slide.


```vb
Sub EffectInfo()

    Dim effIndex As Effect
    Dim seqMain As Sequence

    Set seqMain = ActivePresentation.Slides(1).TimeLine.MainSequence

    For Each effIndex In seqMain
        With effIndex
            MsgBox "Effect Name: " &; .DisplayName &; vbLf &; _
                "Effect Index: " &; .Index
        End With
    Next

End Sub
```


## See also


#### Concepts



[Effect Object](effect-object-powerpoint.md)

