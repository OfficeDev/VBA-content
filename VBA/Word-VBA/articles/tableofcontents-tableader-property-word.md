---
title: TableOfContents.TabLeader Property (Word)
keywords: vbawd10.chm152240138
f1_keywords:
- vbawd10.chm152240138
ms.prod: word
api_name:
- Word.TableOfContents.TabLeader
ms.assetid: aba91b67-33c6-fe8c-0a13-4d5692256091
ms.date: 06/08/2017
---


# TableOfContents.TabLeader Property (Word)

Returns or sets the character between entries and their page numbers in an index, table of authorities, table of contents, or table of figures. Read/write  **[WdTabLeader](wdtableader-enumeration-word.md)** .


## Syntax

 _expression_ . **TabLeader**

 _expression_ Required. A variable that represents a **[TableOfContents](tableofcontents-object-word.md)** collection.


## Example

This example formats the tables of contents in Sales.doc to use a dotted tab leader.


```vb
For Each aTOC In Documents("Sales.doc").TablesOfContents 
 aTOC.TabLeader = wdTabLeaderDots 
Next aTOC
```


## See also


#### Concepts


[TableOfContents Object](tableofcontents-object-word.md)

