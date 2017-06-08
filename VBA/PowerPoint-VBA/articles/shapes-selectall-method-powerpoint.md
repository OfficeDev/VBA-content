---
title: Shapes.SelectAll Method (PowerPoint)
keywords: vbapp10.chm543016
f1_keywords:
- vbapp10.chm543016
ms.prod: powerpoint
api_name:
- PowerPoint.Shapes.SelectAll
ms.assetid: 9d3f5b93-2a8b-5b9a-d725-729baa190a38
ms.date: 06/08/2017
---


# Shapes.SelectAll Method (PowerPoint)

Selects all the shapes in a  **[Shapes](shapes-object-powerpoint.md)** collection.


## Syntax

 _expression_. **SelectAll**

 _expression_ A variable that represents a **Shapes** object.


## Example

This example selects all the shapes on myDocument.


```vb
Set myDocument = ActivePresentation.Slides(1) 
myDocument.Shapes.SelectAll
```


## See also


#### Concepts


[Shapes Object](shapes-object-powerpoint.md)

