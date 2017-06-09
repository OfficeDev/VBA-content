---
title: Presentation.ColorSchemes Property (PowerPoint)
keywords: vbapp10.chm583013
f1_keywords:
- vbapp10.chm583013
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.ColorSchemes
ms.assetid: 4782ee52-3bdd-4459-56da-609a92816692
ms.date: 06/08/2017
---


# Presentation.ColorSchemes Property (PowerPoint)

Returns a  **[ColorSchemes](colorschemes-object-powerpoint.md)** collection that represents the color schemes in the specified presentation. Read-only.


## Syntax

 _expression_. **ColorSchemes**

 _expression_ A variable that represents a **Presentation** object.


### Return Value

ColorSchemes


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


[Presentation Object](presentation-object-powerpoint.md)

