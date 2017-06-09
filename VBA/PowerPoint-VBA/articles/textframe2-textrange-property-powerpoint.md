---
title: TextFrame2.TextRange Property (PowerPoint)
keywords: vbapp10.chm678016
f1_keywords:
- vbapp10.chm678016
ms.prod: powerpoint
api_name:
- PowerPoint.TextFrame2.TextRange
ms.assetid: 288c1209-d12d-fd7c-bc1a-6775d844ca6b
ms.date: 06/08/2017
---


# TextFrame2.TextRange Property (PowerPoint)

Returns a  **[TextRange2 Object (PowerPoint)](textrange2-object-powerpoint.md)** object that represents the text in the specified text frame. Read-only.


## Syntax

 _expression_. **TextRange2**

 _expression_ An expression that returns a **TextFrame2** object.


### Return Value

TextRange2


## Example

This example shows how to set the text for shape one on slide one of the active presentation to the word "Hello!"


```vb
Public Sub TextRange_Example()



    Dim pptSlide As Slide

    Set pptSlide = ActivePresentation.Slides(1)

    pptSlide.Shapes(1).TextFrame2.TextRange = "Hello!"



End Sub
```


## See also


#### Concepts


[TextFrame2 Object](textframe2-object-powerpoint.md)

