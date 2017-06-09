---
title: TimeLine.InteractiveSequences Property (PowerPoint)
keywords: vbapp10.chm649004
f1_keywords:
- vbapp10.chm649004
ms.prod: powerpoint
api_name:
- PowerPoint.TimeLine.InteractiveSequences
ms.assetid: 6dbd6b26-6715-e66c-747f-12f1a16416c8
ms.date: 06/08/2017
---


# TimeLine.InteractiveSequences Property (PowerPoint)

Returns a  **[Sequences](sequences-object-powerpoint.md)** object that represents animations that are triggered by click a specified shape.


## Syntax

 _expression_. **InteractiveSequences**

 _expression_ A variable that represents an **TimeLine** object.


### Return Value

Sequences


## Remarks

The default value of the  **InteractiveSequences** property is an empty **[Sequences](sequences-object-powerpoint.md)** collection.


## Example

The following example adds an interactive sequence to the first slide and sets the text effect properties for the new animation sequence.


```vb
Sub NewInteractiveSeqence()

    Dim seqInteractive As Sequence
    Dim shpText As Shape
    Dim effText As Effect

    Set seqInteractive = ActivePresentation.Slides(1).TimeLine _
        .InteractiveSequences.Add(1)

    Set shpText = ActivePresentation.Slides(1).Shapes(1)
    Set effText = ActivePresentation.Slides(1).TimeLine _
        .MainSequence.AddEffect(Shape:=shpText, _
        EffectId:=msoAnimEffectChangeFont, _
        Trigger:=msoAnimTriggerOnPageClick)

    effText.EffectParameters.FontName = "Broadway"
    seqInteractive.ConvertToTextUnitEffect Effect:=effText, _
        UnitEffect:=msoAnimTextUnitEffectByWord

End Sub
```


## See also


#### Concepts


[TimeLine Object](timeline-object-powerpoint.md)

