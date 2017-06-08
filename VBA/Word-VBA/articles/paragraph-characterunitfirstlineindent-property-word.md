---
title: Paragraph.CharacterUnitFirstLineIndent Property (Word)
keywords: vbawd10.chm156696704
f1_keywords:
- vbawd10.chm156696704
ms.prod: word
api_name:
- Word.Paragraph.CharacterUnitFirstLineIndent
ms.assetid: 013bb6e1-4327-fa20-184f-cde8726e0171
ms.date: 06/08/2017
---


# Paragraph.CharacterUnitFirstLineIndent Property (Word)

Returns or sets the value (in characters) for a first-line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent. Read/write  **Single** .


## Syntax

 _expression_ . **CharacterUnitFirstLineIndent**

 _expression_ Required. A variable that represents a **[Paragraph](paragraph-object-word.md)** object.


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


[Paragraph Object](paragraph-object-word.md)

