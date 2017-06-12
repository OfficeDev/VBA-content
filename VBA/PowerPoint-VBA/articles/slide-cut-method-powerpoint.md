---
title: Slide.Cut Method (PowerPoint)
keywords: vbapp10.chm531012
f1_keywords:
- vbapp10.chm531012
ms.prod: powerpoint
api_name:
- PowerPoint.Slide.Cut
ms.assetid: 03029017-52c8-5176-a218-8b5ff8edec10
ms.date: 06/08/2017
---


# Slide.Cut Method (PowerPoint)

Deletes the specified object and places it on the Clipboard.


## Syntax

 _expression_. **Cut**

 _expression_ A variable that represents a **Slide** object.


## Example

This example deletes slide one from the active presentation and places a copy of it on the Clipboard.


```vb
ActivePresentation.Slides(1).Cut
```


## See also


#### Concepts


[Slide Object](slide-object-powerpoint.md)

