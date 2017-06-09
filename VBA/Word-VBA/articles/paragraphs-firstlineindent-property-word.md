---
title: Paragraphs.FirstLineIndent Property (Word)
keywords: vbawd10.chm156762220
f1_keywords:
- vbawd10.chm156762220
ms.prod: word
api_name:
- Word.Paragraphs.FirstLineIndent
ms.assetid: e882f2da-dc5f-a96d-e18c-39335bd95540
ms.date: 06/08/2017
---


# Paragraphs.FirstLineIndent Property (Word)

Returns or sets the value (in points) for a first line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent. Read/write  **Single** .


## Syntax

 _expression_ . **FirstLineIndent**

 _expression_ A variable that represents a **[Paragraphs](paragraphs-object-word.md)** collection.


## Example

This example sets a first-line indent of 1 inch for the first paragraph in the active document.


```vb
ActiveDocument.Paragraphs.FirstLineIndent = _ 
 InchesToPoints(1)
```

This example sets a hanging indent of 0.5 inch for the second paragraph in the active document. The InchesToPoints method is used to convert inches to points.




```vb
ActiveDocument.Paragraphs.FirstLineIndent = _ 
 InchesToPoints(-0.5)
```


## See also


#### Concepts


[Paragraphs Collection Object](paragraphs-object-word.md)

