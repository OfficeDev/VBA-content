---
title: Document.DeleteAllInkAnnotations Method (Word)
keywords: vbawd10.chm158007775
f1_keywords:
- vbawd10.chm158007775
ms.prod: word
api_name:
- Word.Document.DeleteAllInkAnnotations
ms.assetid: d8446194-f86c-cb48-00e0-82ac84f9bb88
ms.date: 06/08/2017
---


# Document.DeleteAllInkAnnotations Method (Word)

Deletes all handwritten ink annotations in a document.


## Syntax

 _expression_ . **DeleteAllInkAnnotations**

 _expression_ Required. A variable that represents a **[Document](document-object-word.md)** object.


## Remarks

To work with ink annotations, you must be running Microsoft Word on a tablet computer.


## Example

The following example deletes all handwritten ink annotations in the active document.


```vb
ActiveDocument.DeleteAllInkAnnotations
```


## See also


#### Concepts


[Document Object](document-object-word.md)

