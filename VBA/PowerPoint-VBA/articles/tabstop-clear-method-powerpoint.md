---
title: TabStop.Clear Method (PowerPoint)
keywords: vbapp10.chm574005
f1_keywords:
- vbapp10.chm574005
ms.prod: powerpoint
api_name:
- PowerPoint.TabStop.Clear
ms.assetid: bf1bcae7-96a0-6d81-ff7d-806270d95695
ms.date: 06/08/2017
---


# TabStop.Clear Method (PowerPoint)

Clears the specified tab stop from the ruler and deletes it from the  **TabStops** collection.


## Syntax

 _expression_. **Clear**

 _expression_ A variable that represents a **TabStop** object.


## Example

This example clears all tab stops for the text in shape two on slide one in the active presentation.


```vb
With Application.ActivePresentation.Slides(1).Shapes(2).TextFrame _
    .Ruler.TabStops
    For i = .Count To 1 Step -1
        .Item(i).Clear
    Next
End With
```


## See also


#### Concepts


[TabStop Object](tabstop-object-powerpoint.md)

