---
title: ShapeRange.Top Property (PowerPoint)
keywords: vbapp10.chm548037
f1_keywords:
- vbapp10.chm548037
ms.prod: powerpoint
api_name:
- PowerPoint.ShapeRange.Top
ms.assetid: 448b4c64-6519-ce0d-fb2e-9dbc65462494
ms.date: 06/08/2017
---


# ShapeRange.Top Property (PowerPoint)

Returns or sets a  **Single** that represents the distance from the top edge of the topmost shape in the shape range to the top edge of the document. Read/write.


## Syntax

 _expression_. **Top**

 _expression_ A variable that represents a **ShapeRange** object.


### Return Value

Single


## Example

This example arranges windows one and two horizontally; in other words, each window occupies half the available vertical space and all the available horizontal space in the application window's client area. For this example to work, there must be only two document windows open.


```
Windows.Arrange ppArrangeTiled

sngHeight = Windows(1).Height                     ' available height

sngWidth = Windows(1).Width + Windows(2).Width    ' available width

With Windows(1)

    .Width = sngWidth

    .Height = sngHeight / 2

    .Left = 0

End With

With Windows(2)

    .Width = sngWidth

    .Height = sngHeight / 2

    .Top = sngHeight / 2

    .Left = 0

End With
```


## See also


#### Concepts


[ShapeRange Object](shaperange-object-powerpoint.md)

