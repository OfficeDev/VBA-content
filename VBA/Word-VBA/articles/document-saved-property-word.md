---
title: Document.Saved Property (Word)
keywords: vbawd10.chm158007336
f1_keywords:
- vbawd10.chm158007336
ms.prod: word
api_name:
- Word.Document.Saved
ms.assetid: 45bfc77d-2f8e-078c-57c1-ed3ae9f15932
ms.date: 06/08/2017
---


# Document.Saved Property (Word)

 **True** if the specified document or template has not changed since it was last saved. **False** if Microsoft Word displays a prompt to save changes when the document is closed. Read/write **Boolean** .


## Syntax

 _expression_ . **Saved**

 _expression_ A variable that represents a **[Document](document-object-word.md)** object.


## Example

This example saves the active document if it contains previously unsaved changes.


```vb
If ActiveDocument.Saved = False Then ActiveDocument.Save
```


## See also


#### Concepts


[Document Object](document-object-word.md)

