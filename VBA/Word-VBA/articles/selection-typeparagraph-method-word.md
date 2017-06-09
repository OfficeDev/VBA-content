---
title: Selection.TypeParagraph Method (Word)
keywords: vbawd10.chm158663168
f1_keywords:
- vbawd10.chm158663168
ms.prod: word
api_name:
- Word.Selection.TypeParagraph
ms.assetid: e866733b-4800-8e2c-7026-4e9603ccf585
ms.date: 06/08/2017
---


# Selection.TypeParagraph Method (Word)

Inserts a new, blank paragraph.


## Syntax

 _expression_ . **TypeParagraph**

 _expression_ Required. A variable that represents a **[Selection](selection-object-word.md)** object.


## Remarks

This method corresponds to the functionality of the ENTER key. If the selection isn't collapsed to an insertion point, the new paragraph replaces the selection.

Use the  **InsertParagraphAfter** or **InsertParagraphBefore** method to insert a new paragraph without deleting the contents of the selection.


## Example

This example collapses the selection to its end and then inserts a new paragraph following it.


```vb
With Selection 
 .Collapse Direction:=wdCollapseEnd 
 .TypeParagraph 
End With
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

