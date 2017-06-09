---
title: TextRange.AddPeriods Method (PowerPoint)
keywords: vbapp10.chm569032
f1_keywords:
- vbapp10.chm569032
ms.prod: powerpoint
api_name:
- PowerPoint.TextRange.AddPeriods
ms.assetid: 597592ba-6c26-7645-33b8-19ccb375a098
ms.date: 06/08/2017
---


# TextRange.AddPeriods Method (PowerPoint)

Adds a period at the end of each paragraph in the specified text.


## Syntax

 _expression_. **AddPeriods**




## Remarks

This method doesn't add another period at the end of a paragraph that already ends with a period.


## Example

This example adds a period at the end of each paragraph in shape two on slide one in the active presentation.


```vb
Application.ActivePresentation.Slides(1).Shapes(2).TextFrame.TextRange.AddPeriods
```


## See also


#### Concepts


[TextRange Object](textrange-object-powerpoint.md)

