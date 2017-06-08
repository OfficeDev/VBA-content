---
title: Effect.TextRangeStart Property (PowerPoint)
keywords: vbapp10.chm652012
f1_keywords:
- vbapp10.chm652012
ms.prod: powerpoint
api_name:
- PowerPoint.Effect.TextRangeStart
ms.assetid: b6da1565-84e2-acc4-4a06-166c5fda7071
ms.date: 06/08/2017
---


# Effect.TextRangeStart Property (PowerPoint)

Returns or sets the start of a text range. Read-only.


## Syntax

 _expression_. **TextRangeStart**

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

