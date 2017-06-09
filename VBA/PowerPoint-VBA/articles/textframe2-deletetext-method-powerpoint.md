---
title: TextFrame2.DeleteText Method (PowerPoint)
keywords: vbapp10.chm678019
f1_keywords:
- vbapp10.chm678019
ms.prod: powerpoint
api_name:
- PowerPoint.TextFrame2.DeleteText
ms.assetid: 47197c75-99be-4f42-0b4a-bf9207480a94
ms.date: 06/08/2017
---


# TextFrame2.DeleteText Method (PowerPoint)

Deletes the text from a text frame and all the associated properties of the text, including font attributes.


## Syntax

 _expression_. **DeleteText**

 _expression_ An expression that returns a **TextFrame2** object.


### Return Value

Nothing


## Example

This example shows how to delete the text from shape one on slide one of the active presentation, if that shape contains text.


```vb
Public Sub DeleteText_Example()



    Dim pptSlide As Slide

    Set pptSlide = ActivePresentation.Slides(1)

    pptSlide.Shapes(1).TextFrame2.DeleteText



End Sub
```


## See also


#### Concepts


[TextFrame2 Object](textframe2-object-powerpoint.md)

