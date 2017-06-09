---
title: Paragraphs.CharacterUnitLeftIndent Property (Word)
keywords: vbawd10.chm156762239
f1_keywords:
- vbawd10.chm156762239
ms.prod: word
api_name:
- Word.Paragraphs.CharacterUnitLeftIndent
ms.assetid: 692fd810-c3c4-0013-5f16-867105943970
ms.date: 06/08/2017
---


# Paragraphs.CharacterUnitLeftIndent Property (Word)

Returns or sets the left indent value (in characters) for the specified paragraphs. Read/write  **Single** .


## Syntax

 _expression_ . **CharacterUnitLeftIndent**

 _expression_ Required. A variable that represents a **[Paragraphs](paragraphs-object-word.md)** collection.


## Example

This example sets the left indent of the first paragraph in the active document to one character from the left margin.


```vb
ActiveDocument.Paragraphs(1) _ 
 .CharacterUnitLeftIndent = 1
```


## See also


#### Concepts


[Paragraphs Collection Object](paragraphs-object-word.md)

