---
title: Presentation.ExtraColors Property (PowerPoint)
keywords: vbapp10.chm583014
f1_keywords:
- vbapp10.chm583014
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.ExtraColors
ms.assetid: c6a9d155-206c-36e6-c180-aaff8bd85a99
ms.date: 06/08/2017
---


# Presentation.ExtraColors Property (PowerPoint)

Returns an  **[ExtraColors](extracolors-object-powerpoint.md)** object that represents the extra colors available in the specified presentation. Read-only.


## Syntax

 _expression_. **ExtraColors**

 _expression_ A variable that represents an **Presentation** object.


### Return Value

ExtraColors


## Example

The following example adds a rectangle to slide one in the active presentation and sets its fill foreground color to the first extra color. If there hasn't been at least one extra color defined for the presentation, this example will fail.


```vb
With ActivePresentation
    Set rect = .Slides(1).Shapes _
        .AddShape(msoShapeRectangle, 50, 50, 100, 200)
    rect.Fill.ForeColor.RGB = .ExtraColors(1)
End With
```


## See also


#### Concepts


[Presentation Object](presentation-object-powerpoint.md)

