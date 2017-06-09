---
title: TextRange.Cut Method (PowerPoint)
keywords: vbapp10.chm569027
f1_keywords:
- vbapp10.chm569027
ms.prod: powerpoint
api_name:
- PowerPoint.TextRange.Cut
ms.assetid: 9be71668-1486-0466-f87b-47792d402102
ms.date: 06/08/2017
---


# TextRange.Cut Method (PowerPoint)

Deletes the specified object and places it on the Clipboard.


## Syntax

 _expression_. **Cut**

 _expression_ A variable that represents a **TextRange** object.


## Example

This example deletes the text in shape one on slide one in the active presentation and places a copy of it on the Clipboard.


```vb
ActivePresentation.Slides(1).Shapes(1).TextFrame.TextRange.Cut
```


## See also


#### Concepts


[TextRange Object](textrange-object-powerpoint.md)

