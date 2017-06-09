---
title: Document.HasVBProject Property (Word)
keywords: vbawd10.chm158007844
f1_keywords:
- vbawd10.chm158007844
ms.prod: word
api_name:
- Word.Document.HasVBProject
ms.assetid: 1338623e-5832-b77a-cf72-f09d7c8c80de
ms.date: 06/08/2017
---


# Document.HasVBProject Property (Word)

Returns a  **Boolean** that represents whether a document has an attached Microsoft Visual Basic for Applications project. Read-only.


## Syntax

 _expression_ . **HasVBProject**

 _expression_ An expression that returns a **Document** object.


## Remarks

This property is most useful in programatically determining whether a document needs to be saved into a macro-enabled file format. If saved in another format, macros and code projects contained within the document may be lost.


## See also


#### Concepts


[Document Object](document-object-word.md)

