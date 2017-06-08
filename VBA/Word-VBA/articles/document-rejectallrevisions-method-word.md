---
title: Document.RejectAllRevisions Method (Word)
keywords: vbawd10.chm158007614
f1_keywords:
- vbawd10.chm158007614
ms.prod: word
api_name:
- Word.Document.RejectAllRevisions
ms.assetid: d0cf9e63-0057-c832-90b5-e4057c888528
ms.date: 06/08/2017
---


# Document.RejectAllRevisions Method (Word)

Rejects all tracked changes in the specified document.


## Syntax

 _expression_ . **RejectAllRevisions**

 _expression_ Required. A variable that represents a **[Document](document-object-word.md)** object.


## Example

This example checks the main story in active document for tracked changes, and if there are any, the example rejects all revisions in all stories in the document.


```vb
If ActiveDocument.Revisions.Count >= 1 Then _ 
 ActiveDocument.RejectAllRevisions
```


## See also


#### Concepts


[Document Object](document-object-word.md)

