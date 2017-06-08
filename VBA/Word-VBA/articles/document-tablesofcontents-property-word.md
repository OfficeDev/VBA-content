---
title: Document.TablesOfContents Property (Word)
keywords: vbawd10.chm158007327
f1_keywords:
- vbawd10.chm158007327
ms.prod: word
api_name:
- Word.Document.TablesOfContents
ms.assetid: 8c9e923d-c363-281f-d287-3501b980804e
ms.date: 06/08/2017
---


# Document.TablesOfContents Property (Word)

Returns a  **[TablesOfContents](tablesofcontents-object-word.md)** collection that represents the tables of contents in the specified document. Read-only.


## Syntax

 _expression_ . **TablesOfContents**

 _expression_ A variable that represents a **[Document](document-object-word.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example adds a table of contents at the beginning of Sales.doc. The table of contents collects entry text from TC fields.


```vb
Set myRange = Documents("Sales.doc").Range(Start:=0, End:=0) 
Documents("Sales.doc").TablesOfContents.Add Range:=myRange, _ 
 UseFields:=True, UseHeadingStyles:=False
```

This example updates the page numbers for items in the table of contents in the active document.




```vb
For Each myTOC In ActiveDocument.TablesOfContents 
 myTOC.UpdatePageNumbers 
Next myTOC
```


## See also


#### Concepts


[Document Object](document-object-word.md)

