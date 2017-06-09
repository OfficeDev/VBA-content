---
title: PageSetup.SlideHeight Property (PowerPoint)
keywords: vbapp10.chm527004
f1_keywords:
- vbapp10.chm527004
ms.prod: powerpoint
api_name:
- PowerPoint.PageSetup.SlideHeight
ms.assetid: 64b269cf-4b78-eabf-8963-d1971dc90637
ms.date: 06/08/2017
---


# PageSetup.SlideHeight Property (PowerPoint)

Returns or sets the slide height, in points. Read/write.


## Syntax

 _expression_. **SlideHeight**

 _expression_ A variable that represents a **PageSetup** object.


### Return Value

Single


## Example

This example sets the slide height to 8.5 inches and the slide width to 11 inches for the active presentation.


```vb
With Application.ActivePresentation.PageSetup

    .SlideWidth = 11 * 72

    .SlideHeight = 8.5 * 72

End With
```


## See also


#### Concepts


[PageSetup Object](pagesetup-object-powerpoint.md)

