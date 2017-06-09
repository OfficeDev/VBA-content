---
title: Document.IsSubdocument Property (Word)
keywords: vbawd10.chm158007354
f1_keywords:
- vbawd10.chm158007354
ms.prod: word
api_name:
- Word.Document.IsSubdocument
ms.assetid: 2b7bcae0-4934-7563-34e2-d5c5ee6deaeb
ms.date: 06/08/2017
---


# Document.IsSubdocument Property (Word)

 **True** if the specified document is a subdocument of a master document. Read-only **Boolean** .


## Syntax

 _expression_ . **IsSubdocument**

 _expression_ A variable that represents a **[Document](document-object-word.md)** object.


## Example

This example determines whether Sales.doc is a subdocument and then displays a message indicating the status of Sales.doc.


```vb
If Documents("Sales.doc").IsSubdocument = True Then 
 MsgBox "Sales.doc is a subdocument." 
Else 
 MsgBox "Sales.doc is not a subdocument." 
End If
```


## See also


#### Concepts


[Document Object](document-object-word.md)

