---
title: Document.Paragraphs Property (Word)
keywords: vbawd10.chm158007312
f1_keywords:
- vbawd10.chm158007312
ms.prod: word
api_name:
- Word.Document.Paragraphs
ms.assetid: ad60de6b-6287-8ea0-142e-8795f623aa29
ms.date: 06/08/2017
---


# Document.Paragraphs Property (Word)

Returns a  **Paragraphs** collection that represents all the paragraphs in the specified document. Read-only.


## Syntax

 _expression_ . **Paragraphs**

 _expression_ A variable that represents a **[Document](document-object-word.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example sets the line spacing to single for the collection of all paragraphs in section one in the active document.


```vb
ActiveDocument.Sections(1).Range.Paragraphs.LineSpacingRule = _ 
 wdLineSpaceSingle
```

This example sets the line spacing to double for the first paragraph in the selection.




```
Selection.Paragraphs(1).LineSpacingRule = wdLineSpaceDouble
```


## See also


#### Concepts


[Document Object](document-object-word.md)

