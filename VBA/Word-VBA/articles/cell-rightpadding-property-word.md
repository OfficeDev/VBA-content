---
title: Cell.RightPadding Property (Word)
keywords: vbawd10.chm156106866
f1_keywords:
- vbawd10.chm156106866
ms.prod: word
api_name:
- Word.Cell.RightPadding
ms.assetid: 6e71d162-7a8a-9ff2-38ec-c7867804d28b
ms.date: 06/08/2017
---


# Cell.RightPadding Property (Word)

Returns or sets the amount of space (in points) to add to the right of the contents of a single cell or all the cells in a table. Read/write  **Single** .


## Syntax

 _expression_ . **RightPadding**

 _expression_ A variable that represents a **[Cell](cell-object-word.md)** object.


## Remarks

The setting of the  **RightPadding** property for a single cell overrides the setting of the **RightPadding** property for the entire table.


## Example

This example sets the right padding for the first cell in the first row in the first table in the active document to 40 pixels.


```vb
ActiveDocument.Tables(1).Rows(1).Cells(1).RightPadding = _ 
 PixelsToPoints(40, False)
```


## See also


#### Concepts


[Cell Object](cell-object-word.md)

