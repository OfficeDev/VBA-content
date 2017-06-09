---
title: ThemeVariant.Name Property (PowerPoint)
ms.assetid: c28ccf47-05ab-9d75-e190-47a3032faac6
ms.date: 06/08/2017
ms.prod: powerpoint
---


# ThemeVariant.Name Property (PowerPoint)

Returns the name of the theme variation. Read-only.


## Syntax

 _expression_. **Name**

 _expression_ A variable that represents a **ThemeVariant** object.


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


## Property value

 **STRING**


