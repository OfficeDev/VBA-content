---
title: InlineShape.ConvertToShape Method (Word)
keywords: vbawd10.chm162005096
f1_keywords:
- vbawd10.chm162005096
ms.prod: word
api_name:
- Word.InlineShape.ConvertToShape
ms.assetid: 374aea2c-8ff5-d017-4b46-957fafd1bc0a
ms.date: 06/08/2017
---


# InlineShape.ConvertToShape Method (Word)

Converts an inline shape to a free-floating shape. Returns a  **[Shape](shape-object-word.md)** object that represents the new shape.


## Syntax

 _expression_ . **ConvertToShape**

 _expression_ Required. A variable that represents an **[InlineShape](inlineshape-object-word.md)** object.


## Remarks

You must apply the  **AddNodes** method to a **FreeformBuilder** object at least once before you use the **ConvertToShape** method.


## Example

This example converts the first inline shape in the active document to a floating shape.


```vb
ActiveDocument.InlineShapes(1).ConvertToShape
```


## See also


#### Concepts


[InlineShape Object](inlineshape-object-word.md)

