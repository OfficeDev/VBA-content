---
title: ParagraphFormat.CharacterUnitLeftIndent Property (Word)
keywords: vbawd10.chm156434559
f1_keywords:
- vbawd10.chm156434559
ms.prod: word
api_name:
- Word.ParagraphFormat.CharacterUnitLeftIndent
ms.assetid: b54132c9-3d4a-a8d5-2778-c01928f5dda5
ms.date: 06/08/2017
---


# ParagraphFormat.CharacterUnitLeftIndent Property (Word)

Returns or sets the left indent value (in characters) for the specified paragraphs. Read/write  **Single** .


## Syntax

 _expression_ . **CharacterUnitLeftIndent**

 _expression_ Required. A variable that represents a **[ParagraphFormat](paragraphformat-object-word.md)** object.


## Example

This example sets the left indent of the first paragraph in the active document to one character from the left margin.


```vb
ActiveDocument.Paragraphs(1) _ 
 .CharacterUnitLeftIndent = 1
```


## See also


#### Concepts


[ParagraphFormat Object](paragraphformat-object-word.md)

