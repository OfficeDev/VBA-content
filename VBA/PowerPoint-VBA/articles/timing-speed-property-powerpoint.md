---
title: Timing.Speed Property (PowerPoint)
keywords: vbapp10.chm653009
f1_keywords:
- vbapp10.chm653009
ms.prod: powerpoint
api_name:
- PowerPoint.Timing.Speed
ms.assetid: 4dcd7907-47f6-211f-0d88-cfe20165e09f
ms.date: 06/08/2017
---


# Timing.Speed Property (PowerPoint)

Returns or sets the speed, in seconds, of the specified animation. Read/write.


## Syntax

 _expression_. **Speed**

 _expression_ A variable that represents a **Timing** object.


### Return Value

Single


## Example

This example sets the animation for the main sequence to reverse and sets the speed to one second.


```vb
Sub AnimPoints()

    Dim tmlAnim As TimeLine

    Dim spdAnim As Timing



    Set tmlAnim = ActivePresentation.Slides(1).TimeLine

    Set spdAnim = tlnAnim.MainSequence(1).Timing

    With spdAnim

        .AutoReverse = msoTrue

        .Speed = 1

    End With

End Sub
```


## See also


#### Concepts


[Timing Object](timing-object-powerpoint.md)

