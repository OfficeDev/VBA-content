
# SlideRange.ApplyTemplate2 Method (PowerPoint)

 **Last modified:** July 28, 2015

Applies a design template and theme variant to the slide range.

## Syntax

 _expression_. **ApplyTemplate2**(FileName,Variant)

 _expression_A variable that represents a  **SlideRange** object.


### Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
|FileName|Required| **String**|Specifies the name of the design template.|
|Variant|Required| **String**|Specifies the name of the variant to apply.|
|FileName|Required|STRING||
|Variant|Required|STRING||

### Return value

 **VOID**


## Example

This example opens a theme file, gets the ID of the second variant in the theme, and applies it to the slides in the presentation.


```

Sub ChangeThemeVariant()

    Dim name As String
    Dim path As String
    Dim variantID As String
    Dim pptSlideRange As SlideRange
    
    ' Get the name of the active theme family.
    name = ActivePresentation.TemplateName

    ' You need access to the Theme Family in order to access the variants.
    path = "C:\Program Files (x86)\Microsoft Office\Document Themes 15\" &amp; _
        ActivePresentation.TemplateName &amp; ".thmx"

    ' Get the variant ID of the second Variant
    variantID = PowerPoint.Application.OpenThemeFile(path).ThemeVariants(2).Id

    ' Apply that variant to the range of slides.
    Set pptSlideRange = ActivePresentation.Slides.Range
    pptSlideRange.ApplyTemplate2 path, variantID

End Sub
```

