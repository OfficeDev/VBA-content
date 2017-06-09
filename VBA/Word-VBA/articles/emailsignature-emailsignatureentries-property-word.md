---
title: EmailSignature.EmailSignatureEntries Property (Word)
keywords: vbawd10.chm165412969
f1_keywords:
- vbawd10.chm165412969
ms.prod: word
api_name:
- Word.EmailSignature.EmailSignatureEntries
ms.assetid: 8b5a2f6a-d9fe-5f92-d93d-a59e67ee7100
ms.date: 06/08/2017
---


# EmailSignature.EmailSignatureEntries Property (Word)

Returns an  **[EmailSignatureEntries](emailsignatureentries-object-word.md)** object that represents the e-mail signature entries in Microsoft Word. Read-only.


## Syntax

 _expression_ . **EmailSignatureEntries**

 _expression_ An expression that returns an **[EmailSignature](emailsignature-object-word.md)** object.


## Remarks

An e-mail signature is standard text that ends an e-mail message, such as your name and telephone number. Use the  **EmailSignatureEntries** property to create and manage a collection of e-mail signatures that Word will use when creating e-mail messages.


## Example

This example creates a new signature entry based on the author's name and the selection in the active document.


```vb
Sub NewSignature() 
 Application.EmailOptions.EmailSignature _ 
 .EmailSignatureEntries.Add _ 
 Name:=ActiveDocument.BuiltInDocumentProperties("Author"), _ 
 Range:=Selection.Range 
End Sub
```


## See also


#### Concepts


[EmailSignature Object](emailsignature-object-word.md)

