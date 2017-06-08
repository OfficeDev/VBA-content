---
title: Cell.TopPadding Property (Word)
keywords: vbawd10.chm156106863
f1_keywords:
- vbawd10.chm156106863
ms.prod: word
api_name:
- Word.Cell.TopPadding
ms.assetid: 03c8bd07-dde2-6ad3-1291-7b0c0ada424a
ms.date: 06/08/2017
---


# Cell.TopPadding Property (Word)

Returns or sets the amount of space (in points) to add above the contents of a single cell or all the cells in a table. Read/write  **Single** .


## Syntax

 _expression_ . **TopPadding**

 _expression_ A variable that represents a **[Cell](cell-object-word.md)** object.


## Remarks

The setting of the  **TopPadding** property for a single cell overrides the setting of the **TopPadding** property for the entire table.


## Example

This example sets the top padding for the first table in the active document to 40 pixels.


```vb
ActiveDocument.Tables(1).TopPadding = _ 
 PixelsToPoints(40, True)
```


## See also


#### Concepts


[Cell Object](cell-object-word.md)

