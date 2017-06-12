---
title: Document.GetLetterContent Method (Word)
keywords: vbawd10.chm158007420
f1_keywords:
- vbawd10.chm158007420
ms.prod: word
api_name:
- Word.Document.GetLetterContent
ms.assetid: ab0d9fa4-b193-6a7f-641d-d6f971b37457
ms.date: 06/08/2017
---


# Document.GetLetterContent Method (Word)

Retrieves letter elements from the specified document and returns a  **[LetterContent](lettercontent-object-word.md)** object.


## Syntax

 _expression_ . **GetLetterContent**

 _expression_ Required. A variable that represents a **[Document](document-object-word.md)** object.


### Return Value

LetterContent


## Example

This example displays the salutation and recipient name from the letter in the active document.


```vb
MsgBox ActiveDocument.GetLetterContent.Salutation _ 
 &; ActiveDocument.GetLetterContent.RecipientName
```

This example retrieves letter elements from the active document, changes the list of carbon copy (CC) recipients by setting the CClist property, and then uses the SetLetterContent method to update the active document to reflect the changes.




```vb
Set myLetterContent = ActiveDocument.GetLetterContent 
With myLetterContent 
 .CCList = "J. Burns, L. Scarpaczyk, K. Wong" 
 .RecipientName = "Amy Anderson" 
 .RecipientAddress = "123 Main" &; vbCr &; "Bellevue, WA 98004" 
 .LetterStyle = wdFullBlock 
End With 
ActiveDocument.SetLetterContent LetterContent:=myLetterContent
```


## See also


#### Concepts


[Document Object](document-object-word.md)

