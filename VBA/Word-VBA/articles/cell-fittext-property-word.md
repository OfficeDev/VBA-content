---
title: Cell.FitText Property (Word)
keywords: vbawd10.chm156106862
f1_keywords:
- vbawd10.chm156106862
ms.prod: word
api_name:
- Word.Cell.FitText
ms.assetid: ba600e01-1892-557d-95e8-fc9cdea8ef6b
ms.date: 06/08/2017
---


# Cell.FitText Property (Word)

 **True** if Microsoft Word visually reduces the size of text typed into a cell so that it fits within the column width. Read/write **Boolean** .


## Syntax

 _expression_ . **FitText**

 _expression_ A variable that represents a **[Cell](cell-object-word.md)** object.


## Remarks

If the  **FitText** property is set to **True** , the font size of the text is not changed, but the visual width of the characters is adjusted to fit all the typed text into the cell.


## Example

This example sets the first cell in the selection to automatically fit typed text within its width.


```vb
Selection.Cells(1).FitText = True
```


## See also


#### Concepts


[Cell Object](cell-object-word.md)

