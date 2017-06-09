---
title: SlideRange.ColorScheme Property (PowerPoint)
keywords: vbapp10.chm532006
f1_keywords:
- vbapp10.chm532006
ms.prod: powerpoint
api_name:
- PowerPoint.SlideRange.ColorScheme
ms.assetid: 6ae228d8-a105-5745-f7ce-a858bb0954e5
ms.date: 06/08/2017
---


# SlideRange.ColorScheme Property (PowerPoint)

Returns or sets the  **[ColorScheme](colorscheme-object-powerpoint.md)** object that represents the scheme colors for the specified slide, slide range, or slide master. Read/write.


## Syntax

 _expression_. **ColorScheme**

 _expression_ A variable that represents a **SlideRange** object.


### Return Value

ColorScheme


## Example

This example sets the title color to green for slides one and three in the active presentation.


```vb
Set mySlides = ActivePresentation.Slides.Range(Array(1, 3))

mySlides.ColorScheme.Colors(ppTitle).RGB = RGB(0, 255, 0)
```


## See also


#### Concepts


[SlideRange Object](sliderange-object-powerpoint.md)

