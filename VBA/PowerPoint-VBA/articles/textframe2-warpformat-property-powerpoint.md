---
title: TextFrame2.WarpFormat Property (PowerPoint)
keywords: vbapp10.chm678010
f1_keywords:
- vbapp10.chm678010
ms.prod: powerpoint
api_name:
- PowerPoint.TextFrame2.WarpFormat
ms.assetid: 1b22dbf3-d54f-7a00-46b1-6dd1b84b0993
ms.date: 06/08/2017
---


# TextFrame2.WarpFormat Property (PowerPoint)

Returns or sets the warp format (how the text is warped) for the specified text frame. Read/write.


## Syntax

 _expression_. **WarpFormat**

 _expression_ An expression that returns a **TextFrame2** object.


### Return Value

MsoWarpFormat


## Remarks

The value of the  **WarpFormat** property can be one of the **[MsoWarpFormat](http://msdn.microsoft.com/library/481cead3-900f-66b6-8200-21342b0ce21c%28Office.15%29.aspx)** constants.


## Example

This example shows how to set the warp format for shape one on slide one of the active presentation.


```vb
Public Sub WarpFormat_Example() 
 
    Dim pptSlide As Slide 
    Set pptSlide = ActivePresentation.Slides(1) 
    pptSlide.Shapes(1).TextFrame2.WarpFormat = msoWarpFormat15 
     
End Sub
```


## See also


#### Concepts


[TextFrame2 Object](textframe2-object-powerpoint.md)

