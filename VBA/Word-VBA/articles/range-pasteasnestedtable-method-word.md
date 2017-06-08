---
title: Range.PasteAsNestedTable Method (Word)
keywords: vbawd10.chm157155550
f1_keywords:
- vbawd10.chm157155550
ms.prod: word
api_name:
- Word.Range.PasteAsNestedTable
ms.assetid: 8d7a3fc6-5fc2-9cbc-d551-b4606af54619
ms.date: 06/08/2017
---


# Range.PasteAsNestedTable Method (Word)

Pastes a cell or group of cells as a nested table into the selected range.


## Syntax

 _expression_ . **PasteAsNestedTable**

 _expression_ Required. A variable that represents a **[Range](range-object-word.md)** object.


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


[Range Object](range-object-word.md)

