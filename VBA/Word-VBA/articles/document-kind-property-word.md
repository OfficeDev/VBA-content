---
title: Document.Kind Property (Word)
keywords: vbawd10.chm158007339
f1_keywords:
- vbawd10.chm158007339
ms.prod: word
api_name:
- Word.Document.Kind
ms.assetid: 2a2ca204-ae61-4de2-feaa-678f564b2ca0
ms.date: 06/08/2017
---


# Document.Kind Property (Word)

Returns or sets the format type that Microsoft Word uses when automatically formatting the specified document. Read/write  **[WdDocumentKind](wddocumentkind-enumeration-word.md)** .


## Syntax

 _expression_ . **Kind**

 _expression_ Required. A variable that represents a **[Document](document-object-word.md)** object.


## Example

This example asks the user whether the active document is an e-mail message. If the response is Yes, the document is formatted as an e-mail message.


```vb
response = MsgBox("Is this document an e-mail message?", vbYesNo) 
If response = vbYes Then 
 ActiveDocument.Kind = wdDocumentEmail 
 ActiveDocument.Content.AutoFormat 
End If
```


## See also


#### Concepts


[Document Object](document-object-word.md)

