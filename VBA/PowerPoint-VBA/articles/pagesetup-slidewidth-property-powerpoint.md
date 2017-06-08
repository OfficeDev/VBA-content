---
title: PageSetup.SlideWidth Property (PowerPoint)
keywords: vbapp10.chm527005
f1_keywords:
- vbapp10.chm527005
ms.prod: powerpoint
api_name:
- PowerPoint.PageSetup.SlideWidth
ms.assetid: 671d3962-a4d0-fcca-009e-784abaedbd8f
ms.date: 06/08/2017
---


# PageSetup.SlideWidth Property (PowerPoint)

Returns or sets the slide width, in points. Read/write.


## Syntax

 _expression_. **SlideWidth**

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

