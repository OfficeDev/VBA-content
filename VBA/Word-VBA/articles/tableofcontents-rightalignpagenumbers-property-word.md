---
title: TableOfContents.RightAlignPageNumbers Property (Word)
keywords: vbawd10.chm152240135
f1_keywords:
- vbawd10.chm152240135
ms.prod: word
api_name:
- Word.TableOfContents.RightAlignPageNumbers
ms.assetid: f14e4b13-a6d4-0085-af31-ef4077b5104f
ms.date: 06/08/2017
---


# TableOfContents.RightAlignPageNumbers Property (Word)

 **True** if page numbers are aligned with the right margin in a table of contents. Read/write **Boolean** .


## Syntax

 _expression_ . **RightAlignPageNumbers**

 _expression_ Required. A variable that represents a **[TableOfContents](tableofcontents-object-word.md)** collection.


## Example

This example right-aligns page numbers for the first table of contents in the active document.


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

