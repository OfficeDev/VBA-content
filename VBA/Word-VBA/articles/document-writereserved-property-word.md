---
title: Document.WriteReserved Property (Word)
keywords: vbawd10.chm158007384
f1_keywords:
- vbawd10.chm158007384
ms.prod: word
api_name:
- Word.Document.WriteReserved
ms.assetid: be5d8696-9e72-f8a3-2b47-a2fde13359f9
ms.date: 06/08/2017
---


# Document.WriteReserved Property (Word)

 **True** if the specified document is protected with a write password. Read-only **Boolean** .


## Syntax

 _expression_ . **WriteReserved**

 _expression_ A variable that represents a **[Document](document-object-word.md)** object.


## Example

This example displays a message if the active document has a write password.


```vb
If ActiveDocument.WriteReserved = True Then 
 MsgBox "Changes cannot be made to this document." 
End If
```


## See also


#### Concepts


[Document Object](document-object-word.md)

