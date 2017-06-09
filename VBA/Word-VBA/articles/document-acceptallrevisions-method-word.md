---
title: Document.AcceptAllRevisions Method (Word)
keywords: vbawd10.chm158007613
f1_keywords:
- vbawd10.chm158007613
ms.prod: word
api_name:
- Word.Document.AcceptAllRevisions
ms.assetid: 3281313c-fa16-1f68-0435-f822f7cea06d
ms.date: 06/08/2017
---


# Document.AcceptAllRevisions Method (Word)

Accepts all tracked changes in the specified document.


## Syntax

 _expression_ . **AcceptAllRevisions**

 _expression_ Required. A variable that represents a **[Document](document-object-word.md)** object.


## Example

This example checks the main story in the active document for tracked changes, and if there are any, the example incorporates all revisions in all stories in the document.


```vb
If ActiveDocument.Revisions.Count >= 1 Then _ 
 ActiveDocument.AcceptAllRevisions
```


## See also


#### Concepts


[Document Object](document-object-word.md)

