---
title: LetterContent.Closing Property (Word)
keywords: vbawd10.chm161546361
f1_keywords:
- vbawd10.chm161546361
ms.prod: word
api_name:
- Word.LetterContent.Closing
ms.assetid: 46f367a8-c487-a837-f37c-7570e005728d
ms.date: 06/08/2017
---


# LetterContent.Closing Property (Word)

Returns or sets the closing text for a letter created by the Letter Wizard (for example, "Sincerely yours"). Read/write  **String** .


## Syntax

 _expression_ . **Closing**

 _expression_ A variable that represents a **[LetterContent](lettercontent-object-word.md)** object.


## Example

This example displays the closing text from the active document.


```vb
MsgBox ActiveDocument.GetLetterContent.Closing
```

This example retrieves letter elements from the active document, changes the closing text by setting the  **Closing** property, and then uses the **[SetLetterContent](document-setlettercontent-method-word.md)** method to update the document to reflect the changes.




```vb
Dim lcCurrent As LetterContent 
 
Set lcCurrent = ActiveDocument.GetLetterContent 
lcCurrent.Closing = "Sincerely yours," 
ActiveDocument.SetLetterContent LetterContent:=lcCurrent
```


## See also


#### Concepts


[LetterContent Object](lettercontent-object-word.md)

