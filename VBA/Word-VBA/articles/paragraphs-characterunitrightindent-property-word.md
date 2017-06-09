---
title: Paragraphs.CharacterUnitRightIndent Property (Word)
keywords: vbawd10.chm156762238
f1_keywords:
- vbawd10.chm156762238
ms.prod: word
api_name:
- Word.Paragraphs.CharacterUnitRightIndent
ms.assetid: dbbb903b-924b-1f36-3e56-9489f544f601
ms.date: 06/08/2017
---


# Paragraphs.CharacterUnitRightIndent Property (Word)

Returns or sets the right indent value (in characters) for the specified paragraphs. Read/write  **Single** .


## Syntax

 _expression_ . **CharacterUnitRightIndent**

 _expression_ Required. A variable that represents a **[Paragraphs](paragraphs-object-word.md)** collection.


## Example

This example sets the right indent for all paragraphs in the active document to one character from the right margin.


```vb
ActiveDocument.Paragraphs _ 
 .CharacterUnitRightIndent = 1
```


## See also


#### Concepts


[Paragraphs Collection Object](paragraphs-object-word.md)

