---
title: Paragraph.CharacterUnitLeftIndent Property (Word)
keywords: vbawd10.chm156696703
f1_keywords:
- vbawd10.chm156696703
ms.prod: word
api_name:
- Word.Paragraph.CharacterUnitLeftIndent
ms.assetid: 1dbe6053-52fd-f17c-aa95-3cfdef1222d5
ms.date: 06/08/2017
---


# Paragraph.CharacterUnitLeftIndent Property (Word)

Returns or sets the left indent value (in characters) for the specified paragraphs. Read/write  **Single** .


## Syntax

 _expression_ . **CharacterUnitLeftIndent**

 _expression_ Required. A variable that represents a **[Paragraph](paragraph-object-word.md)** object.


## Example

This example sets the left indent of the first paragraph in the active document to one character from the left margin.


```vb
ActiveDocument.Paragraphs(1) _ 
 .CharacterUnitLeftIndent = 1
```


## See also


#### Concepts


[Paragraph Object](paragraph-object-word.md)

