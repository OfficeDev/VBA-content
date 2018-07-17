---
title: ExtraColors Object (PowerPoint)
keywords: vbapp10.chm529000
f1_keywords:
- vbapp10.chm529000
ms.prod: powerpoint
api_name:
- PowerPoint.ExtraColors
ms.assetid: 8f13d8cd-be83-21d6-ebd1-855d9289a65e
ms.date: 06/08/2017
---


# ExtraColors Object (PowerPoint)

Represents the extra colors in a presentation. The object can contain up to eight colors, each of which is represented by an red-green-blue (RGB) value.


## Example

Use the [ExtraColors](presentation-extracolors-property-powerpoint.md)property to return the  **ExtraColors** object. Use **ExtraColors** (index), where index is the extra color index number, to return the red-green-blue (RGB) value for a single extra color. The following example adds a rectangle to slide one in the active presentation and sets its fill foreground color to the first extra color. If there hasn't been at least one extra color defined for the presentation, this example will fail.


```vb
With ActivePresentation
    Set rect = .Slides(1).Shapes _
        .AddShape(msoShapeRectangle, 50, 50, 100, 200)
    rect.Fill.ForeColor.RGB = .ExtraColors(1)
End With
```

Use the [Add](extracolors-add-method-powerpoint.md)method to add an extra color. The following example adds an extra color to the active presentation (if the color hasn't already been added).




```vb
ActivePresentation.ExtraColors.Add RGB(69, 32, 155)
```


## See also


#### Concepts


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)

