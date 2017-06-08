---
title: ColorSchemes Object (PowerPoint)
keywords: vbapp10.chm536000
f1_keywords:
- vbapp10.chm536000
ms.prod: powerpoint
api_name:
- PowerPoint.ColorSchemes
ms.assetid: 9b062448-88f5-b38d-2c76-330c691c9d72
ms.date: 06/08/2017
---


# ColorSchemes Object (PowerPoint)

A collection of all the  **[ColorScheme](colorscheme-object-powerpoint.md)** objects in the specified presentation. Each **ColorScheme** object represents a color scheme, which is a set of colors that are used together on a slide.


## Example

Use the [ColorSchemes](presentation-colorschemes-property-powerpoint.md)property to return the  **ColorSchemes** collection. Use **ColorSchemes** (index), where index is the color scheme index number, to return a single **ColorScheme** object. The following example deletes color scheme two from the active presentation.


```vb
ActivePresentation.ColorSchemes(2).Delete
```

Use the [Add](colorschemes-add-method-powerpoint.md)method to create a new color scheme and add it to the  **ColorSchemes** collection. The following example adds a color scheme to the active presentation and sets the title color and background color for the color scheme (because no argument was used with the **Add** method, the added color scheme is initially identical to the first standard color scheme in the presentation).




```vb
With ActivePresentation.ColorSchemes.Add

    .Colors(ppTitle).RGB = RGB(255, 0, 0)

    .Colors(ppBackground).RGB = RGB(128, 128, 0)

End With
```

Set the [ColorScheme](slide-colorscheme-property-powerpoint.md)property of a  **Slide**, **SlideRange**, or **Master** object to return the color scheme for one slide, a set of slides, or a master, respectively. The following example sets the color scheme for all the slides in the active presentation to the third color scheme in the presentation.




```vb
With ActivePresentation

    .Slides.Range.ColorScheme = .ColorSchemes(3)

End With
```


## See also


#### Concepts


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)

