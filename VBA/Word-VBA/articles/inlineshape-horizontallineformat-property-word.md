---
title: InlineShape.HorizontalLineFormat Property (Word)
keywords: vbawd10.chm162005111
f1_keywords:
- vbawd10.chm162005111
ms.prod: word
api_name:
- Word.InlineShape.HorizontalLineFormat
ms.assetid: 3e6f3887-d906-a761-d1ee-a4c4560c4888
ms.date: 06/08/2017
---


# InlineShape.HorizontalLineFormat Property (Word)

Returns a  **[HorizontalLineFormat](horizontallineformat-object-word.md)** object that contains the horizontal line formatting for the specified **InlineShape** object. Read-only.


## Syntax

 _expression_ . **HorizontalLineFormat**

 _expression_ A variable that represents a **[InlineShape](inlineshape-object-word.md)** object.


## Example

This example sets the length of the specified horizontal line to 50% of the window width.


```vb
ActiveDocument.InlineShapes(1).HorizontalLineFormat _ 
 .PercentWidth = 50
```


## See also


#### Concepts


[InlineShape Object](inlineshape-object-word.md)

