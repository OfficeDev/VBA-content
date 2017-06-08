---
title: TableOfContents.IncludePageNumbers Property (Word)
keywords: vbawd10.chm152240136
f1_keywords:
- vbawd10.chm152240136
ms.prod: word
api_name:
- Word.TableOfContents.IncludePageNumbers
ms.assetid: 2097f009-ae18-70c3-3f3b-dbabcee06fa5
ms.date: 06/08/2017
---


# TableOfContents.IncludePageNumbers Property (Word)

 **True** if page numbers are included in the table of contents. Read/write **Boolean** .


## Syntax

 _expression_ . **IncludePageNumbers**

 _expression_ Required. A variable that represents a **[TableOfContents](tableofcontents-object-word.md)** collection.


## Example

This example formats the first table of contents in the active document to include right-aligned page numbers.


```vb
If ActiveDocument.TablesOfContents.Count >= 1 Then 
 With ActiveDocument.TablesOfContents(1) 
 .IncludePageNumbers = True 
 .RightAlignPageNumbers = True 
 End With 
End If
```


## See also


#### Concepts


[TableOfContents Object](tableofcontents-object-word.md)

