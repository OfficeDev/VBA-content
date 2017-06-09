---
title: TableOfContents.TableID Property (Word)
keywords: vbawd10.chm152240133
f1_keywords:
- vbawd10.chm152240133
ms.prod: word
api_name:
- Word.TableOfContents.TableID
ms.assetid: d95186f5-b6ee-20cd-840e-e55ec3f06d04
ms.date: 06/08/2017
---


# TableOfContents.TableID Property (Word)

Returns or sets a one-letter identifier that's used to build a table of contents from TOC fields. Read/write  **String** .


## Syntax

 _expression_ . **TableID**

 _expression_ Required. A variable that represents a **[TableOfContents](tableofcontents-object-word.md)** collection.


## Remarks

This property corresponds to the \f switch for a TOC field. For example, "T" builds a table of contents from TOC fields using the table identifier T.


## Example

This example inserts a TOC field with an "A" identifier after the selection in Sales.doc. The first table of contents is then formatted to compile "A" entries.


```vb
Documents("Sales.doc").TablesOfContents.MarkEntry _ 
 Range:=Selection.Range, _ 
 Entry:="Hello", TableID:="A" 
With Documents("Sales.doc").TablesOfContents(1) 
 .TableID = "A" 
 .UseFields = True 
 .UseHeadingStyles = False 
 .Update 
End With
```


## See also


#### Concepts


[TableOfContents Object](tableofcontents-object-word.md)

