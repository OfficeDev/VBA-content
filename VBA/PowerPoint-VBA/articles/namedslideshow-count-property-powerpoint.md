---
title: NamedSlideShow.Count Property (PowerPoint)
keywords: vbapp10.chm516006
f1_keywords:
- vbapp10.chm516006
ms.prod: powerpoint
api_name:
- PowerPoint.NamedSlideShow.Count
ms.assetid: 09aeed71-dfc6-2ee6-1430-c5e7f0ed2bc1
ms.date: 06/08/2017
---


# NamedSlideShow.Count Property (PowerPoint)

Returns the number of objects in the specified collection. Read-only.


## Syntax

 _expression_. **Count**

 _expression_ A variable that represents a **NamedSlideShow** object.


### Return Value

Long


## Example

This example closes all windows except the active window.


```vb
With Application.Windows

    For i = 2 To .Count

        .Item(2).Close

    Next

End With
```


## See also


#### Concepts


[NamedSlideShow Object](namedslideshow-object-powerpoint.md)

