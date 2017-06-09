---
title: ParagraphFormat.CharacterUnitRightIndent Property (Word)
keywords: vbawd10.chm156434558
f1_keywords:
- vbawd10.chm156434558
ms.prod: word
api_name:
- Word.ParagraphFormat.CharacterUnitRightIndent
ms.assetid: ef9476ce-fa19-3c75-934b-9dd33da23076
ms.date: 06/08/2017
---


# ParagraphFormat.CharacterUnitRightIndent Property (Word)

Returns or sets the right indent value (in characters) for the specified paragraphs. Read/write  **Single** .


## Syntax

 _expression_ . **CharacterUnitRightIndent**

 _expression_ Required. A variable that represents a **[ParagraphFormat](paragraphformat-object-word.md)** object.


## Example

This example sets the right indent for all paragraphs in the active document to one character from the right margin.


```vb
ActiveDocument.Paragraphs _ 
 .CharacterUnitRightIndent = 1
```


## See also


#### Concepts


[ParagraphFormat Object](paragraphformat-object-word.md)

