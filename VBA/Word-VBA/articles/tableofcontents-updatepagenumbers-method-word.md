---
title: TableOfContents.UpdatePageNumbers Method (Word)
keywords: vbawd10.chm152240229
f1_keywords:
- vbawd10.chm152240229
ms.prod: word
api_name:
- Word.TableOfContents.UpdatePageNumbers
ms.assetid: 3b7e3080-c2bb-0a4b-2062-f1a774eeb715
ms.date: 06/08/2017
---


# TableOfContents.UpdatePageNumbers Method (Word)

Updates the page numbers for items in the specified table of contents.


## Syntax

 _expression_ . **UpdatePageNumbers**

 _expression_ Required. A variable that represents a **[TableOfContents](tableofcontents-object-word.md)** collection.


## Example

This example inserts a page break at the insertion point and then updates the page numbers for the first table of contents in the active document.


```
Selection.Collapse Direction:=wdCollapseStart 
Selection.InsertBreak Type:=wdPageBreak 
ActiveDocument.TablesOfContents(1).UpdatePageNumbers
```


## See also


#### Concepts


[TableOfContents Object](tableofcontents-object-word.md)

