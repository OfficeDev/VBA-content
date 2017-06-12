---
title: Paragraphs.CharacterUnitFirstLineIndent Property (Word)
keywords: vbawd10.chm156762240
f1_keywords:
- vbawd10.chm156762240
ms.prod: word
api_name:
- Word.Paragraphs.CharacterUnitFirstLineIndent
ms.assetid: 0d11652c-1617-1975-0b1d-e07284966e90
ms.date: 06/08/2017
---


# Paragraphs.CharacterUnitFirstLineIndent Property (Word)

Returns or sets the value (in characters) for a first-line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent. Read/write  **Single** .


## Syntax

 _expression_ . **CharacterUnitFirstLineIndent**

 _expression_ Required. A variable that represents a **[Paragraphs](paragraphs-object-word.md)** collection.


## Example

This example sets a first-line indent of one character for the first paragraph in the active document.


```vb
ActiveDocument.Paragraphs(1) _ 
 .CharacterUnitFirstLineIndent = 1
```

This example sets a hanging indent of 1.5 characters for the second paragraph in the active document.




```vb
ActiveDocument.Paragraphs(2) _ 
 .CharacterUnitFirstLineIndent = -1.5
```


## See also


#### Concepts


[Paragraphs Collection Object](paragraphs-object-word.md)

