---
title: SlideRange.ApplyTemplate2 Method (PowerPoint)
keywords: vbapp10.chm532048
f1_keywords:
- vbapp10.chm532048
ms.assetid: e8d0dcae-d311-a9bd-4cf9-ac171ec15cbb
ms.date: 06/08/2017
ms.prod: powerpoint
---


# SlideRange.ApplyTemplate2 Method (PowerPoint)

Applies a design template and theme variant to the slide range.


## Syntax

 _expression_. **ApplyTemplate2**_(FileName,_ _Variant)_

 _expression_ A variable that represents a **SlideRange** object.


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

This example opens a theme file, gets the ID of the second variant in the theme, and applies it to the slides in the presentation.


```vb
Sub ChangeThemeVariant()

    Dim name As String
    Dim path As String
    Dim variantID As String
    Dim pptSlideRange As SlideRange
    
    ' Get the name of the active theme family.
    name = ActivePresentation.TemplateName

    ' You need access to the Theme Family in order to access the variants.
    path = "C:\Program Files (x86)\Microsoft Office\Document Themes 15\" &; _
        ActivePresentation.TemplateName &; ".thmx"

    ' Get the variant ID of the second Variant
    variantID = PowerPoint.Application.OpenThemeFile(path).ThemeVariants(2).Id

    ' Apply that variant to the range of slides.
    Set pptSlideRange = ActivePresentation.Slides.Range
    pptSlideRange.ApplyTemplate2 path, variantID

End Sub
```


