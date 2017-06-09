---
title: Document.FormFields Property (Word)
keywords: vbawd10.chm158007317
f1_keywords:
- vbawd10.chm158007317
ms.prod: word
api_name:
- Word.Document.FormFields
ms.assetid: ed97fd75-0da5-b008-26c6-ea16465fddc1
ms.date: 06/08/2017
---


# Document.FormFields Property (Word)

Returns a  **[FormFields](formfields-object-word.md)** collection that represents all the form fields in the document. Read-only.


## Syntax

 _expression_ . **FormFields**

 _expression_ A variable that represents a **[Document](document-object-word.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example sets the content of the form field named "Text1" to "Name."


```vb
ActiveDocument.FormFields("Text1").Result = "Name"
```


## See also


#### Concepts


[Document Object](document-object-word.md)

