---
title: Slide.ApplyTemplate2 Method (PowerPoint)
keywords: vbapp10.chm531044
f1_keywords:
- vbapp10.chm531044
ms.assetid: e4931f7b-98de-a854-3752-c1f9ca70cf3b
ms.date: 06/08/2017
ms.prod: powerpoint
---


# Slide.ApplyTemplate2 Method (PowerPoint)

Applies a design template and theme variant to the slide.


## Syntax

 _expression_. **ApplyTemplate2**_(FileName,_ _Variant)_

 _expression_ A variable that represents a **Slide** object.


### Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FileName_|Required|**String**|Specifies the name of the design template.|
| _Variant_|Required|**String**|Specifies the name of the variant to apply.|
| _FileName_|Required|STRING||
| _Variant_|Required|STRING||
| _VariantGUID_|Required|STRING||

### Return value

 **VOID**


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


