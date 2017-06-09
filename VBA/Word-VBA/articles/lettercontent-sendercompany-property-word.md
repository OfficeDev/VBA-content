---
title: LetterContent.SenderCompany Property (Word)
keywords: vbawd10.chm161546362
f1_keywords:
- vbawd10.chm161546362
ms.prod: word
api_name:
- Word.LetterContent.SenderCompany
ms.assetid: 7f4abf0c-baf8-bb63-6e9e-58360a3b019b
ms.date: 06/08/2017
---


# LetterContent.SenderCompany Property (Word)

Returns or sets the company name of the person creating a letter with the Letter Wizard. Read/write  **String** .


## Syntax

 _expression_ . **SenderCompany**

 _expression_ An expression that returns a **[LetterContent](lettercontent-object-word.md)** object.


## Example

This example retrieves the Letter Wizard elements from the active document. If the sender's company name isn't blank, the example displays the text in a message box.


```vb
If ActiveDocument.GetLetterContent.SenderCompany <> "" Then 
 MsgBox ActiveDocument.GetLetterContent.SenderCompany 
End If
```


## See also


#### Concepts


[LetterContent Object](lettercontent-object-word.md)

