---
title: ThemeVariant Object (PowerPoint)
keywords: vbapp10.chm738000
f1_keywords:
- vbapp10.chm738000
ms.assetid: de00374f-05fd-4cae-08f8-ef417cd944b5
ms.date: 06/08/2017
ms.prod: powerpoint
---


# ThemeVariant Object (PowerPoint)

Represents a variation (set of colors and fonts) in a theme.


## Example

This example opens a theme file, iterates through the variants in the theme, and prints the name and ID of each variation in the theme.


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
    ' its name and ID.
    For Each pptThemeVariant In pptThemeVariants
    
        Debug.Print "Variation " &; pptThemeVariant.name &; " id: " &; pptThemeVariant.Id
    
    Next pptThemeVariant

End Sub
```


