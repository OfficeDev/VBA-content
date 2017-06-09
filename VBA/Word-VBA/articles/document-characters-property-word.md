---
title: Document.Characters Property (Word)
keywords: vbawd10.chm158007315
f1_keywords:
- vbawd10.chm158007315
ms.prod: word
api_name:
- Word.Document.Characters
ms.assetid: 1703bbe3-6c46-a45b-9f36-1205a0d2d47c
ms.date: 06/08/2017
---


# Document.Characters Property (Word)

Returns a  **[Characters](characters-object-word.md)** collection that represents the characters in a document. Read-only.


## Syntax

 _expression_ . **Characters**

 _expression_ A variable that represents a **[Document](document-object-word.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example returns the number of characters in the first sentence in the active document (spaces are included in the count).


```
numchars = ActiveDocument.Characters.Count
```


## See also


#### Concepts


[Document Object](document-object-word.md)

