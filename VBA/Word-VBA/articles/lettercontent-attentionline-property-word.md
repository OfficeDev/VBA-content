---
title: LetterContent.AttentionLine Property (Word)
keywords: vbawd10.chm161546355
f1_keywords:
- vbawd10.chm161546355
ms.prod: word
api_name:
- Word.LetterContent.AttentionLine
ms.assetid: 56cbda4c-08ff-2d0b-2b1b-2c5e0ac26fea
ms.date: 06/08/2017
---


# LetterContent.AttentionLine Property (Word)

Returns or sets the attention line text for a letter created by the Letter Wizard. Read/write  **String** .


## Syntax

 _expression_ . **AttentionLine**

 _expression_ A variable that represents a **[LetterContent](lettercontent-object-word.md)** object.


## Example

This example retrieves the Letter Wizard elements from the active document. If the attention line isn't blank, the example displays the text in a message box.


```vb
If ActiveDocument.GetLetterContent.AttentionLine <> "" Then 
 MsgBox ActiveDocument.GetLetterContent.AttentionLine 
End If
```

This example retrieves the Letter Wizard elements from the active document, changes the attention line text, and then uses the  **[SetLetterContent](document-setlettercontent-method-word.md)** method to update the document to reflect the changes.




```vb
Dim lcTemp As LetterContent 
 
Set lcTemp = ActiveDocument.GetLetterContent 
 
lcTemp.AttentionLine = "Greetings" 
ActiveDocument.SetLetterContent LetterContent:=lcTemp
```


## See also


#### Concepts


[LetterContent Object](lettercontent-object-word.md)

