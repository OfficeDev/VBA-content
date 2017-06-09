---
title: Master.ColorScheme Property (PowerPoint)
keywords: vbapp10.chm533005
f1_keywords:
- vbapp10.chm533005
ms.prod: powerpoint
api_name:
- PowerPoint.Master.ColorScheme
ms.assetid: f481aa76-e96f-686a-edbb-b2bef8be0e8c
ms.date: 06/08/2017
---


# Master.ColorScheme Property (PowerPoint)

Returns or sets the  **[ColorScheme](colorscheme-object-powerpoint.md)** object that represents the scheme colors for the specified slide, slide range, or slide master. Read/write.


## Syntax

 _expression_. **ColorScheme**

 _expression_ A variable that represents a **Master** object.


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


[Master Object](master-object-powerpoint.md)

