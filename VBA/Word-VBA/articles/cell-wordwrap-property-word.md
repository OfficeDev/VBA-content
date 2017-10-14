---
title: Cell.WordWrap Property (Word)
keywords: vbawd10.chm156106860
f1_keywords:
- vbawd10.chm156106860
ms.prod: word
api_name:
- Word.Cell.WordWrap
ms.assetid: 16255023-d6c3-3c27-402f-490970b7af33
ms.date: 06/08/2017
---


# Cell.WordWrap Property (Word)

 **True** if Microsoft Word wraps text to multiple lines and lengthens the cell so that the cell width remains the same. Read/write **Boolean** .


## Syntax

 _expression_ . **WordWrap**

 _expression_ Required. A variable that represents a **[Cell](cell-object-word.md)** object.


## Example

This example sets Microsoft Word to wrap text to multiple lines in the third cell of the first table so that the cell's width remains the same.


```vb
ActiveDocument.Tables(1).Range.Cells(3).WordWrap = True
```


## See also


#### Concepts


[Cell Object](cell-object-word.md)

