---
title: LetterContent.RecipientName Property (Word)
keywords: vbawd10.chm161546348
f1_keywords:
- vbawd10.chm161546348
ms.prod: word
api_name:
- Word.LetterContent.RecipientName
ms.assetid: e5e75700-5189-1189-7454-fc74214f5e35
ms.date: 06/08/2017
---


# LetterContent.RecipientName Property (Word)

Returns or sets the name of the person who'll be receiving the letter created by the Letter Wizard. Read/write  **String** .


## Syntax

 _expression_ . **RecipientName**

 _expression_ An expression that returns a **[LetterContent](lettercontent-object-word.md)** object.


## Example

This example displays the salutation and recipient name for the active document.


```vb
MsgBox ActiveDocument.GetLetterContent.Salutation _ 
 &; Space(1) &; ActiveDocument.GetLetterContent.RecipientName
```


## See also


#### Concepts


[LetterContent Object](lettercontent-object-word.md)

