---
title: Document.TablesOfAuthorities Property (Word)
keywords: vbawd10.chm158007328
f1_keywords:
- vbawd10.chm158007328
ms.prod: word
api_name:
- Word.Document.TablesOfAuthorities
ms.assetid: c49d1fc5-1d0a-3b6e-ab9e-62b968766cd3
ms.date: 06/08/2017
---


# Document.TablesOfAuthorities Property (Word)

Returns a  **[TableOfAuthorities](tableofauthorities-object-word.md)** collection that represents the tables of authorities in the specified document. Read-only.


## Syntax

 _expression_ . **TablesOfAuthorities**

 _expression_ A variable that represents a **[Document](document-object-word.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example adds a table of authorities at the beginning of Sales.doc. The table of authorities compiles references from all categories.


```vb
Set myRange = Documents("Sales.doc").Range(Start:=0, End:=0) 
Documents("Sales.doc").TablesOfAuthorities.Add Range:=myRange, _ 
 Category:=0, Passim:=True, IncludeCategoryHeader:=True
```

This example updates each table of authorities in the active document.




```vb
For Each myTOA In ActiveDocument.TablesOfAuthorities 
 myTOA.Update 
Next myTOA
```


## See also


#### Concepts


[Document Object](document-object-word.md)

