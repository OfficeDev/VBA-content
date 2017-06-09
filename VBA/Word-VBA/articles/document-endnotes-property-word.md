---
title: Document.Endnotes Property (Word)
keywords: vbawd10.chm158007304
f1_keywords:
- vbawd10.chm158007304
ms.prod: word
api_name:
- Word.Document.Endnotes
ms.assetid: 3c3e87c0-ea76-8bc4-0b2e-755bff6aa14c
ms.date: 06/08/2017
---


# Document.Endnotes Property (Word)

Returns an  **[Endnotes](endnotes-object-word.md)** collection that represents all the endnotes in a document. Read-only.


## Syntax

 _expression_ . **Endnotes**

 _expression_ A variable that represents a **[Document](document-object-word.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example positions the endnotes in the active document at the end of the document and formats the endnote reference marks as lowercase roman numerals.


```vb
With ActiveDocument.Endnotes 
 .Location = wdEndOfDocument 
 .NumberStyle = wdNoteNumberStyleLowercaseRoman 
End With
```


## See also


#### Concepts


[Document Object](document-object-word.md)

