---
title: Paragraphs.IncreaseSpacing Method (Word)
keywords: vbawd10.chm156762447
f1_keywords:
- vbawd10.chm156762447
ms.prod: word
api_name:
- Word.Paragraphs.IncreaseSpacing
ms.assetid: d0416601-5616-0e93-540f-f09e192b0c91
ms.date: 06/08/2017
---


# Paragraphs.IncreaseSpacing Method (Word)

Increases the spacing before and after paragraphs in six-point increments.


## Syntax

 _expression_ . **IncreaseSpacing**

 _expression_ Required. A variable that represents a **[Paragraphs](paragraphs-object-word.md)** collection.


## Example

This example increases the before and after spacing of a paragraph or selection of paragraphs by six points each time the procedure is run.


```vb
Sub IncreaseParaSpacing() 
 Selection.Paragraphs.IncreaseSpacing 
End Sub
```


## See also


#### Concepts


[Paragraphs Collection Object](paragraphs-object-word.md)

