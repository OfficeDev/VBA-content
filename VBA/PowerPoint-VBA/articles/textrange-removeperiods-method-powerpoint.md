---
title: TextRange.RemovePeriods Method (PowerPoint)
keywords: vbapp10.chm569033
f1_keywords:
- vbapp10.chm569033
ms.prod: powerpoint
api_name:
- PowerPoint.TextRange.RemovePeriods
ms.assetid: 66562cc7-e25b-e110-000e-c01b62caf761
ms.date: 06/08/2017
---


# TextRange.RemovePeriods Method (PowerPoint)

Removes the period at the end of each paragraph in the specified text.


## Syntax

 _expression_. **RemovePeriods**

 _expression_ A variable that represents a **TextRange** object.


## Example

This example removes the period at the end of each paragraph in shape two on slide one in the active presentation.


```vb
Application.ActivePresentation.Slides(1) _
    .Shapes(2).TextFrame.TextRange _
    .RemovePeriods
```


## See also


#### Concepts


[TextRange Object](textrange-object-powerpoint.md)

