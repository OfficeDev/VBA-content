---
title: Application.OpenThemeFile Method (PowerPoint)
keywords: vbapp10.chm502070
f1_keywords:
- vbapp10.chm502070
ms.assetid: b34d5a6f-8cf8-ce6a-3c0c-c1ed43c413c6
ms.date: 06/08/2017
ms.prod: powerpoint
---


# Application.OpenThemeFile Method (PowerPoint)

Opens the specified theme file (*thmx).


## Syntax

 _expression_. **OpenThemeFile**_(themeFileName)_

 _expression_ A variable that represents a **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
| _themeFileName_|Required|**String**|The path of the theme file (*.thmx) to open.|
| _themeFileName_|Required|STRING||

### Return value

[Theme](theme-object-powerpoint.md)


## Example

This example opens a theme file, gets the ID of the third variant in the theme, and applies it to the first slide in the presentation.


```vb
Sub ChangeThemeVariant()

    Dim name As String
    Dim path As String
    Dim variantID As String
    
    ' Get the name of the active theme family.
    name = ActivePresentation.TemplateName

    ' You need access to the Theme Family in order to access the variants.
    path = "C:\Program Files (x86)\Microsoft Office\Document Themes 15\" &; _
        ActivePresentation.TemplateName &; ".thmx"

    ' Get the variant ID of the third Variant
    ' and apply that variant to the presentation.
    variantID = PowerPoint.Application.OpenThemeFile(path).ThemeVariants(3).Id
    ActivePresentation.Slides(1).ApplyTemplate2 path, variantID

End Sub
```


