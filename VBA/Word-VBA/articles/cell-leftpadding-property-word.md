---
title: Cell.LeftPadding Property (Word)
keywords: vbawd10.chm156106865
f1_keywords:
- vbawd10.chm156106865
ms.prod: word
api_name:
- Word.Cell.LeftPadding
ms.assetid: b80dba74-7f12-0258-de03-e9941b6b1f4c
ms.date: 06/08/2017
---


# Cell.LeftPadding Property (Word)

Returns or sets the amount of space (in points) to add to the left of the contents of a single cell or all the cells in a table. Read/write  **Single** .


## Syntax

 _expression_ . **LeftPadding**

 _expression_ A variable that represents a **[Cell](cell-object-word.md)** object.


## Remarks

The setting of the  **LeftPadding** property for a single cell overrides the setting of the **LeftPadding** property for the entire table.


## Example

This example sets the left padding for the first cell in the first table in the active document to 40 pixels.


```vb
ActiveDocument.Tables(1).Rows(1).Cells(1).LeftPadding = _ 
 PixelsToPoints(40, False)
```


## See also


#### Concepts


[Cell Object](cell-object-word.md)

