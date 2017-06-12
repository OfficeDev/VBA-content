---
title: Document.Indexes Property (Word)
keywords: vbawd10.chm158007335
f1_keywords:
- vbawd10.chm158007335
ms.prod: word
api_name:
- Word.Document.Indexes
ms.assetid: 47a8a5d3-3c3c-81f0-8d51-5459c5bc7f89
ms.date: 06/08/2017
---


# Document.Indexes Property (Word)

Returns an  **[Indexes](indexes-object-word.md)** collection that represents all the indexes in the specified document. Read-only.


## Syntax

 _expression_ . **Indexes**

 _expression_ A variable that represents a **[Document](document-object-word.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example adds an index at the end of the active document.


```vb
Set MyRange = _ 
 ActiveDocument.Range(Start:=ActiveDocument.Content.End - 1, _ 
 End:=ActiveDocument.Content.End - 1) 
ActiveDocument.Indexes.Add Range:=MyRange, NumberOfColumns:=1, _ 
 HeadingSeparator:=False
```

This example inserts an index entry for the selected text.




```vb
If Selection.Type = wdSelectionNormal Then 
 ActiveDocument.Indexes.MarkEntry Range:=Selection.Range, _ 
 Entry:=Selection.Range.Text 
End If
```


## See also


#### Concepts


[Document Object](document-object-word.md)

