---
title: Selection.PasteAsNestedTable Method (Word)
keywords: vbawd10.chm158663189
f1_keywords:
- vbawd10.chm158663189
ms.prod: word
api_name:
- Word.Selection.PasteAsNestedTable
ms.assetid: 42a2f604-694e-6b39-23d2-d8c453618222
ms.date: 06/08/2017
---


# Selection.PasteAsNestedTable Method (Word)

Pastes a cell or group of cells as a nested table into the selection.


## Syntax

 _expression_ . **PasteAsNestedTable**

 _expression_ Required. A variable that represents a **[Selection](selection-object-word.md)** object.


## Remarks

You can use  **PasteAsNestedTable** only if the Clipboard contains a cell or group of cells and the selected range is a cell or group of cells in the current document.


## Example

This example pastes the contents of the Clipboard into the third cell of the first table in the active document.


```vb
ActiveDocument.Tables(1).Rows(1).Cells(3).Range _ 
 .PasteAsNestedTable
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

