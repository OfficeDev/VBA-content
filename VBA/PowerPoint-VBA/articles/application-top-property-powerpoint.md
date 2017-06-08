---
title: Application.Top Property (PowerPoint)
keywords: vbapp10.chm502027
f1_keywords:
- vbapp10.chm502027
ms.prod: powerpoint
api_name:
- PowerPoint.Application.Top
ms.assetid: d8c70fc9-e0f1-ed53-7a22-150838599719
ms.date: 06/08/2017
---


# Application.Top Property (PowerPoint)

Returns or sets a  **Single** that represents the distance in points from the top edge of the document, application, and slide show window to the top edge of the application window's client area. Read/write.


## Syntax

 _expression_. **Top**

 _expression_ A variable that represents an **Application** object.


### Return Value

Single


## Remarks

Setting this property to a very large positive or negative value may position the window completely off the desktop.


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


[Application Object](application-object-powerpoint.md)

