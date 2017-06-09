---
title: Document.Revisions Property (Word)
keywords: vbawd10.chm158007326
f1_keywords:
- vbawd10.chm158007326
ms.prod: word
api_name:
- Word.Document.Revisions
ms.assetid: 26211417-b9c5-128e-1b00-cb312dd3724f
ms.date: 06/08/2017
---


# Document.Revisions Property (Word)

Returns a  **Revisions** collection that represents the tracked changes in the document or range. Read-only.


## Syntax

 _expression_ . **Revisions**

 _expression_ A variable that represents a **[Document](document-object-word.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example displays the number of tracked changes in the first section in the active document.


```vb
MsgBox ActiveDocument.Sections(1).Range.Revisions.Count
```

This example accepts all tracked changes in the first paragraph in the selection.




```vb
Set myRange = Selection.Paragraphs(1).Range 
myRange.Revisions.AcceptAll
```


## See also


#### Concepts


[Document Object](document-object-word.md)

