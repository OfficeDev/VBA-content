---
title: Slide.ColorScheme Property (PowerPoint)
keywords: vbapp10.chm531006
f1_keywords:
- vbapp10.chm531006
ms.prod: powerpoint
api_name:
- PowerPoint.Slide.ColorScheme
ms.assetid: 3d40d93f-4e7d-e95f-8340-d138da2a1b55
ms.date: 06/08/2017
---


# Slide.ColorScheme Property (PowerPoint)

Returns or sets the  **[ColorScheme](colorscheme-object-powerpoint.md)** object that represents the scheme colors for the specified slide, slide range, or slide master. Read/write.


## Syntax

 _expression_. **ColorScheme**

 _expression_ A variable that represents a **Slide** object.


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


[Slide Object](slide-object-powerpoint.md)

