---
title: Index.TabLeader Property (Word)
keywords: vbawd10.chm159186950
f1_keywords:
- vbawd10.chm159186950
ms.prod: word
api_name:
- Word.Index.TabLeader
ms.assetid: 82bc6e93-1dd7-aa56-1fca-8fcb9ed72784
ms.date: 06/08/2017
---


# Index.TabLeader Property (Word)

Returns or sets the leader character between entries in an index and their associated page numbers. Read/write  **WdTabLeader** .


## Syntax

 _expression_ . **TabLeader**

 _expression_ Required. A variable that represents an **[Index](index-object-word.md)** object.


## Example

This example adds an index at the end of the active document. The page numbers are right-aligned with a dashed-line tab leader.


```vb
Set myRange = ActiveDocument.Range( _ 
 Start:=ActiveDocument.Content.End -1, _ 
 End:=ActiveDocument.Content.End -1) 
ActiveDocument.Indexes.Add(Range:=myRange, Type:=wdIndexIndent, _ 
 RightAlignPageNumbers:=True).TabLeader = wdTabLeaderDashes
```


## See also


#### Concepts


[Index Object](index-object-word.md)

