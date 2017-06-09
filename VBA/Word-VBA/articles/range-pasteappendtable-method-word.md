---
title: Range.PasteAppendTable Method (Word)
keywords: vbawd10.chm157155742
f1_keywords:
- vbawd10.chm157155742
ms.prod: word
api_name:
- Word.Range.PasteAppendTable
ms.assetid: dc3b9914-b0d6-aa85-a357-a96475680caf
ms.date: 06/08/2017
---


# Range.PasteAppendTable Method (Word)

Merges pasted cells into an existing table by inserting the pasted rows between the selected rows. No cells are overwritten.


## Syntax

 _expression_ . **PasteAppendTable**

 _expression_ Required. A variable that represents a **[Range](range-object-word.md)** object.


## Example

This example pastes table cells by inserting rows into the current table at the insertion point. This example assumes that the Clipboard contains a collection of table cells.


```vb
Sub PasteAppend 
 Selection.PasteAppendTable 
End Sub
```


## See also


#### Concepts


[Range Object](range-object-word.md)

