---
title: Theme.ThemeVariants Property (PowerPoint)
keywords: vbapp10.chm740002
f1_keywords:
- vbapp10.chm740002
ms.assetid: f88d8c2c-4964-7392-5457-25216ece92d0
ms.date: 06/08/2017
ms.prod: powerpoint
---


# Theme.ThemeVariants Property (PowerPoint)

Returns a  **[ThemeVariants](themevariant-object-powerpoint.md)** collection that represents the variations in the theme.


## Syntax

 _expression_. **ThemeVariants**

 _expression_ A variable that represents a **Theme** object.


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


## Property value

 **THEMEVARIANTS**


