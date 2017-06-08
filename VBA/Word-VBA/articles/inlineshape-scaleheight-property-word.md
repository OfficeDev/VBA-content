---
title: InlineShape.ScaleHeight Property (Word)
keywords: vbawd10.chm162005002
f1_keywords:
- vbawd10.chm162005002
ms.prod: word
api_name:
- Word.InlineShape.ScaleHeight
ms.assetid: c8f07ca4-4f0c-c365-1962-4404ca7a6ed4
ms.date: 06/08/2017
---


# InlineShape.ScaleHeight Property (Word)

Scales the height of the specified inline shape relative to its original size. Read/write  **Single** .


## Syntax

 _expression_ . **ScaleHeight**

 _expression_ Required. A variable that represents an **[InlineShape](inlineshape-object-word.md)** object.


## Example

This example sets the height and width of the first inline shape in the active document to 150 percent of the shape's original height and width.


```vb
With ActiveDocument.InlineShapes(1) 
 .ScaleHeight = 150 
 .ScaleWidth = 150 
End With
```


## See also


#### Concepts


[InlineShape Object](inlineshape-object-word.md)

