---
title: ColorFormat.RGB Property (PowerPoint)
keywords: vbapp10.chm506002
f1_keywords:
- vbapp10.chm506002
ms.prod: powerpoint
api_name:
- PowerPoint.ColorFormat.RGB
ms.assetid: 5bb68052-5931-2096-277c-fb44c76b37eb
ms.date: 06/08/2017
---


# ColorFormat.RGB Property (PowerPoint)

Returns or sets the red-green-blue (RGB) value of the specified color. Read/write.


## Syntax

 _expression_. **RGB**

 _expression_ A variable that represents a **ColorFormat** object.


### Return Value

MsoRGBType


## Example

This example sets the background color for color scheme three in the active presentation and then applies the color scheme to all slides in the presentation that are based on the slide master.


```vb
With ActivePresentation

    Set cs1 = .ColorSchemes(3)

    cs1.Colors(ppBackground).RGB = RGB(128, 128, 0)

    .SlideMaster.ColorScheme = cs1

End With
```


## See also


#### Concepts


[ColorFormat Object](colorformat-object-powerpoint.md)

