---
title: Paragraph.FirstLineIndent Property (Word)
keywords: vbawd10.chm156696684
f1_keywords:
- vbawd10.chm156696684
ms.prod: word
api_name:
- Word.Paragraph.FirstLineIndent
ms.assetid: 44f326b6-3352-da1a-5ff0-952627ed7b90
ms.date: 06/08/2017
---


# Paragraph.FirstLineIndent Property (Word)

Returns or sets the value (in points) for a first line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent. Read/write  **Single** .


## Syntax

 _expression_ . **FirstLineIndent**

 _expression_ A variable that represents a **[Paragraph](paragraph-object-word.md)** object.


## Example

This example sets a first-line indent of 1 inch for the first paragraph in the active document.


```vb
ActiveDocument.Paragraphs(1).FirstLineIndent = _ 
 InchesToPoints(1)
```

This example sets a hanging indent of 0.5 inch for the second paragraph in the active document. The InchesToPoints method is used to convert inches to points.




```vb
ActiveDocument.Paragraphs(2).FirstLineIndent = _ 
 InchesToPoints(-0.5)
```


## See also


#### Concepts


[Paragraph Object](paragraph-object-word.md)

