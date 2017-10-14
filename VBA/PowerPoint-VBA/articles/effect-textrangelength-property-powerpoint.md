---
title: Effect.TextRangeLength Property (PowerPoint)
keywords: vbapp10.chm652013
f1_keywords:
- vbapp10.chm652013
ms.prod: powerpoint
api_name:
- PowerPoint.Effect.TextRangeLength
ms.assetid: b68690a5-f93e-0833-73be-a6259d604064
ms.date: 06/08/2017
---


# Effect.TextRangeLength Property (PowerPoint)

Returns or sets a  **Long** that represents the length of a text range. Read-only.


## Syntax

 _expression_. **TextRangeLength**

 _expression_ A variable that represents a **Effect** object.


### Return Value

Long


## Example

This example adds a shape with text and rotates the shape without rotating the text.


```vb
Sub SetTextRange()

    Dim shpStar As Shape
    Dim sldOne As Slide
    Dim effNew As Effect

    Set sldOne = ActivePresentation.Slides(1)
    Set shpStar = sldOne.Shapes.AddShape(Type:=msoShape5pointStar, _
        Left:=32, Top:=32, Width:=300, Height:=300)

    shpStar.TextFrame.TextRange.Text = "Animated shape."
    Set effNew = sldOne.TimeLine.MainSequence.AddEffect(Shape:=shpStar, _
        EffectId:=msoAnimEffectPath5PointStar, Level:=msoAnimateTextByAllLevels, _
        Trigger:=msoAnimTriggerAfterPrevious)

    With effNew
        If .TextRangeStart = 0 And .TextRangeLength > 0 Then
            With .Behaviors.Add(Type:=msoAnimTypeRotation).RotationEffect
                .From = 0
                .To = 360
            End With
            .Timing.AutoReverse = msoTrue
        End If
    End With

End Sub
```


## See also


#### Concepts



[Effect Object](effect-object-powerpoint.md)

