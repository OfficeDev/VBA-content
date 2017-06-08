---
title: Shapes.SelectAll Method (Publisher)
keywords: vbapb10.chm2162726
f1_keywords:
- vbapb10.chm2162726
ms.prod: publisher
api_name:
- Publisher.Shapes.SelectAll
ms.assetid: 67b88529-814d-c029-1bde-e5dade87636a
ms.date: 06/08/2017
---


# Shapes.SelectAll Method (Publisher)

Selects all the shapes in the specified  **[Shapes](shapes-object-publisher.md)** collection.


## Syntax

 _expression_. **SelectAll**

 _expression_A variable that represents a  **Shapes** object.


## Example

This example selects all the shapes on page one of the active publication.


```vb
ActiveDocument.Pages(1).Shapes.SelectAll
```


