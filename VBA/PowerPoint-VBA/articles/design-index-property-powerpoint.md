---
title: Design.Index Property (PowerPoint)
keywords: vbapp10.chm644007
f1_keywords:
- vbapp10.chm644007
ms.prod: powerpoint
api_name:
- PowerPoint.Design.Index
ms.assetid: 16a9ca67-4db4-c7a4-118b-553f0d7efc98
ms.date: 06/08/2017
---


# Design.Index Property (PowerPoint)

Returns a  **Long** that represents the index number for an animation effect or design. Read-only.


## Syntax

 _expression_. **Index**

 _expression_ A variable that represents a **Design** object.


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


[Design Object](design-object-powerpoint.md)

