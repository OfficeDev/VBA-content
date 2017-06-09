---
title: Index.Filter Property (Word)
keywords: vbawd10.chm159186953
f1_keywords:
- vbawd10.chm159186953
ms.prod: word
api_name:
- Word.Index.Filter
ms.assetid: 87b5ad20-cc3d-b1d5-9622-ff23ea25120c
ms.date: 06/08/2017
---


# Index.Filter Property (Word)

Returns or sets a value that specifies how Microsoft Word classifies the first character of entries in the specified index.read/write  **Long** . Can be one of the following **wdIndexFilter** constants.


## Syntax

 _expression_ . **Filter**

 _expression_ A variable that represents an **[Index](index-object-word.md)** object.


## Example

This example inserts an index at the end of the active document. right-aligns the page numbers, and then sets Microsoft Word to classify index entries as "wdIndexFilterAkasatana".


```vb
Set myRange = ActiveDocument.Range _ 
 (Start:=ActiveDocument.Content.End -1, _ 
 End:=ActiveDocument.Content.End -1) 
ActiveDocument.Indexes.Add(Range:=myRange, Type:=wdIndexIndent, _ 
 RightAlignPageNumbers:=True).Filter = wdIndexFilterAkasatana
```


## See also


#### Concepts


[Index Object](index-object-word.md)

