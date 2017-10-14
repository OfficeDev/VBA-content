---
title: TextFrame2.WordArtFormat Property (PowerPoint)
keywords: vbapp10.chm678011
f1_keywords:
- vbapp10.chm678011
ms.prod: powerpoint
api_name:
- PowerPoint.TextFrame2.WordArtFormat
ms.assetid: 7ab4d90b-aae1-d98e-50d2-14b181d370ba
ms.date: 06/08/2017
---


# TextFrame2.WordArtFormat Property (PowerPoint)

Returns or sets the WordArt type for the specified text frame. Read/write.


## Syntax

 _expression_. **WordArtFormat**

 _expression_ An expression that returns a **TextFrame2** object.


### Return Value

MsoPresetTextEffect


## Remarks

The value of the  **WordArtFormat** property can be one of these **[MsoPresetTextEffect](http://msdn.microsoft.com/library/56a7008d-ce2c-f127-56de-851cb8fef44f%28Office.15%29.aspx)** constants.


## Example

This example shows how to set the word art format for shape one on slide one in the active presentation.


```vb
Public Sub WordArtFormat_Example() 
 
    Dim pptSlide As Slide 
    Set pptSlide = ActivePresentation.Slides(1) 
    pptSlide.Shapes(1).TextFrame2.WordArtFormat = msoTextEffect20 
     
End Sub
```


## See also


#### Concepts


[TextFrame2 Object](textframe2-object-powerpoint.md)

