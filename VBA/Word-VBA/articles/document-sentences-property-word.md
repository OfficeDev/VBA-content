---
title: Document.Sentences Property (Word)
keywords: vbawd10.chm158007314
f1_keywords:
- vbawd10.chm158007314
ms.prod: word
api_name:
- Word.Document.Sentences
ms.assetid: 41906136-815c-4dfc-ad92-c16ad420ab91
ms.date: 06/08/2017
---


# Document.Sentences Property (Word)

Returns a  **[Sentences](sentences-object-word.md)** collection that represents all the sentences in the document. Read-only.


## Syntax

 _expression_ . **Sentences**

 _expression_ A variable that represents a **[Document](document-object-word.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example copies the first sentences in the active document.


```vb
ActiveDocument.Sentences(1).Copy
```

This example deletes the last sentence in the active document.




```vb
ActiveDocument.Sentences.Last.Delete
```


## See also


#### Concepts


[Document Object](document-object-word.md)

