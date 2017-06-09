---
title: Document.Footnotes Property (Word)
keywords: vbawd10.chm158007303
f1_keywords:
- vbawd10.chm158007303
ms.prod: word
api_name:
- Word.Document.Footnotes
ms.assetid: 6257f658-69f5-4223-153b-56bc3791a99d
ms.date: 06/08/2017
---


# Document.Footnotes Property (Word)

Returns a  **[Footnotes](footnotes-object-word.md)** collection that represents all the footnotes in a document. Read-only.


## Syntax

 _expression_ . **Footnotes**

 _expression_ A variable that represents a **[Document](document-object-word.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example changes the footnote reference marks for the footnotes in the active document to lowercase letters, starting with the letter "c".


```vb
With ActiveDocument.Footnotes 
 .StartingNumber = 3 
 .NumberStyle = wdNoteNumberStyleLowercaseLetter 
End With
```

This example inserts an automatically numbered footnote at the insertion point.




```
Selection.Collapse Direction:=wdCollapseStart 
ActiveDocument.Footnotes.Add Range:=Selection.Range, _ 
 Text:="(Lone Creek Press, 1995)"
```


## See also


#### Concepts


[Document Object](document-object-word.md)

