---
title: TabStop.Position Property (PowerPoint)
keywords: vbapp10.chm574004
f1_keywords:
- vbapp10.chm574004
ms.prod: powerpoint
api_name:
- PowerPoint.TabStop.Position
ms.assetid: fc7e75a5-e0a3-78de-91d9-b116f1ded321
ms.date: 06/08/2017
---


# TabStop.Position Property (PowerPoint)

Returns or sets the position of the specified tab stop, in points. Read/write.


## Syntax

 _expression_. **Position**

 _expression_ A variable that represents a **TabStop** object.


### Return Value

Single


## Example

This example deletes all tab stops greater than 1 inch (72 points) for the text in shape two on slide one in the active presentation.


```vb
With Application.ActivePresentation.Slides(1).Shapes(2).TextFrame _
    .Ruler.TabStops
    For i = .Count To 1 Step -1
        With .Item(i)
            If .Position > 72 Then .Clear
        End With
    Next
End With
```


## See also


#### Concepts


[TabStop Object](tabstop-object-powerpoint.md)

