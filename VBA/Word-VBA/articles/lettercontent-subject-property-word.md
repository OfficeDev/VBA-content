---
title: LetterContent.Subject Property (Word)
keywords: vbawd10.chm161546356
f1_keywords:
- vbawd10.chm161546356
ms.prod: word
api_name:
- Word.LetterContent.Subject
ms.assetid: cfdf65ed-7a92-6462-b868-74c4cd3b17e2
ms.date: 06/08/2017
---


# LetterContent.Subject Property (Word)

Returns or sets the subject text of a letter created by the Letter Wizard. Read/write  **String** .


## Syntax

 _expression_ . **Subject**

 _expression_ Required. A variable that represents a **[LetterContent](lettercontent-object-word.md)** object.


## Example

This example displays the subject of a letter created by the Letter Wizard, unless the subject is an empty string.


```vb
If ActiveDocument.GetLetterContent.Subject <> "" Then 
 MsgBox ActiveDocument.GetLetterContent.Subject 
End If
```


## See also


#### Concepts


[LetterContent Object](lettercontent-object-word.md)

