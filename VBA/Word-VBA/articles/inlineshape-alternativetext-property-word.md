---
title: InlineShape.AlternativeText Property (Word)
keywords: vbawd10.chm162005123
f1_keywords:
- vbawd10.chm162005123
ms.prod: word
api_name:
- Word.InlineShape.AlternativeText
ms.assetid: a9eba1a4-393d-7d85-a964-984d7b5bd485
ms.date: 06/08/2017
---


# InlineShape.AlternativeText Property (Word)

Returns or sets a  **String** that represents the alternative text associated with a shape in a Web page. Read/write.


## Syntax

 _expression_ . **AlternativeText**

 _expression_ A variable that represents an **[InlineShape](inlineshape-object-word.md)** object.


## Example

The following example sets the alternative text for the selected shape in the active window. The selected shape is a picture of a mallard duck.


```vb
ActiveWindow.Selection.Shapes(1) _ 
 .AlternativeText = "This is a mallard duck."
```


## See also


#### Concepts


[InlineShape Object](inlineshape-object-word.md)

