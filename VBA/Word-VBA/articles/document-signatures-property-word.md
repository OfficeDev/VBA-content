---
title: Document.Signatures Property (Word)
keywords: vbawd10.chm158007635
f1_keywords:
- vbawd10.chm158007635
ms.prod: word
api_name:
- Word.Document.Signatures
ms.assetid: 2f6cf537-6f7a-9cca-1d2c-39bb581630ad
ms.date: 06/08/2017
---


# Document.Signatures Property (Word)

Returns a  **SignatureSet** collection that represents the digital signatures for a document.


## Syntax

 _expression_ . **Signatures**

 _expression_ An expression that returns a **[Document](document-object-word.md)** object.


## Remarks

To digitally sign Microsoft Word documents and verify other signatures in them, you will need the Microsoft CryptoAPI and a unique digital signature certificate. The CryptoAPI is installed with Microsoft Internet Explorer 4.01 and higher. You can obtain a digital signature certificate from a certification authority.


## Example

This example displays the  **Signatures** dialog box with which you can add a digital signature to a document.


```vb
Sub AddSignature 
 ActiveDocument.Signatures.Add 
End Sub
```


## See also


#### Concepts


[Document Object](document-object-word.md)

