---
title: RGBColor Object (PowerPoint)
keywords: vbapp10.chm538000
f1_keywords:
- vbapp10.chm538000
ms.prod: powerpoint
api_name:
- PowerPoint.RGBColor
ms.assetid: 1da5054f-7eaa-37e8-9a5b-d90c790de576
ms.date: 06/08/2017
---


# RGBColor Object (PowerPoint)

Represents a single color in a color scheme.


## Example

Use the [Colors](colorscheme-colors-method-powerpoint.md)method to return an  **RGBColor** object. You can set an **RGBColor** object to another **RGBColor** object. You can use the[RGB](rgbcolor-rgb-property-powerpoint.md)property to set or return the explicit red-green-blue value for an  **RGBColor** object, with the exception of the **RGBColor** objects defined by the **ppNotSchemeColor** and **ppSchemeColorMixed** constants. The **RGB** property can be returned, but not set, for these two objects. The following example sets the background color in color scheme one in the active presentation to red and sets the title color to the title color that's defined for color scheme two.


```vb
With ActivePresentation.ColorSchemes

    .Item(1).Colors(ppBackground).RGB = RGB(255, 0, 0)

    .Item(1).Colors(ppTitle) = .Item(2).Colors(ppTitle)

End With
```


## See also


#### Concepts


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)

