---
title: TextRange.LtrRun Method (PowerPoint)
keywords: vbapp10.chm569039
f1_keywords:
- vbapp10.chm569039
ms.prod: powerpoint
api_name:
- PowerPoint.TextRange.LtrRun
ms.assetid: 5c6787cc-d37c-8aec-b49e-12418291e006
ms.date: 06/08/2017
---


# TextRange.LtrRun Method (PowerPoint)

Sets the direction of text in a text range to read from left to right.


## Syntax

 _expression_. **LtrRun**

 _expression_ A variable that represents a **TextRange** object.


## Remarks

This method makes it possible to use text from both left-to-right and right-to-left languages in the same presentation.


## Example

The following example finds all of the shapes on slide one that contain text and changes the text to read from left to right.


```vb
ActiveWindow.ViewType = ppViewSlide

For Each sh In ActivePresentation.Slides(1).Shapes

    If sh.HasTextFrame Then

         sh.TextFrame.TextRange.LtrRun

    End If

Next
```


## See also


#### Concepts


[TextRange Object](textrange-object-powerpoint.md)

