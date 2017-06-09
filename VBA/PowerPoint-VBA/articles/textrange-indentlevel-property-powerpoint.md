---
title: TextRange.IndentLevel Property (PowerPoint)
keywords: vbapp10.chm569025
f1_keywords:
- vbapp10.chm569025
ms.prod: powerpoint
api_name:
- PowerPoint.TextRange.IndentLevel
ms.assetid: 3ba39fc4-6fc4-62ca-0e87-a7605d6c8bea
ms.date: 06/08/2017
---


# TextRange.IndentLevel Property (PowerPoint)

Returns or sets the the indent level for the specified text as an integer from 1 to 5, where 1 indicates a first-level paragraph with no indentation. Read/write.


## Syntax

 _expression_. **IndentLevel**

 _expression_ A variable that represents an **TextRange** object.


### Return Value

Long


## Example

This example indents the second paragraph in shape two on slide two in the active presentation.


```vb
Application.ActivePresentation.Slides(2).Shapes(2).TextFrame _
    .TextRange.Paragraphs(2).IndentLevel = 2
```


## See also


#### Concepts


[TextRange Object](textrange-object-powerpoint.md)

