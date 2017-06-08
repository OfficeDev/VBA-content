---
title: Theme Object (PowerPoint)
keywords: vbapp10.chm740000
f1_keywords:
- vbapp10.chm740000
ms.assetid: f541387f-6cf4-1bae-97e4-534ef7fba040
ms.date: 06/08/2017
ms.prod: powerpoint
---


# Theme Object (PowerPoint)

Represents a theme (a collection of colors, fonts, and effects).


## Example

The following code example gets a reference to the currently active theme and then iterates over each theme variation in the theme.


```vb
Sub IterateThemeVariants()

    Dim pptTheme As Theme
    Dim pptThemeVariants As ThemeVariants
    Dim pptThemeVariant As ThemeVariant
    Dim path As String
    
    ' Get a reference to the currently active theme.
    path = "C:\Program Files (x86)\Microsoft Office\Document Themes 15\" &; _
        ActivePresentation.TemplateName &; ".thmx"
    Set pptTheme = Application.OpenThemeFile(path)
    
    ' Get a reference to all of the variations in the theme.
    Set pptThemeVariants = pptTheme.ThemeVariants
    
    ' Iterate over each variation of the theme and print
    ' its ID.
    For Each pptThemeVariant In pptThemeVariants
    
        Debug.Print "Variation id: " &; pptThemeVariant.Id
    
    Next pptThemeVariant

End Sub
```


