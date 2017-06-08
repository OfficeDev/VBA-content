---
title: Cell.BottomPadding Property (Word)
keywords: vbawd10.chm156106864
f1_keywords:
- vbawd10.chm156106864
ms.prod: word
api_name:
- Word.Cell.BottomPadding
ms.assetid: 5f265dc2-a9c4-d307-69a8-1f73407a4301
ms.date: 06/08/2017
---


# Cell.BottomPadding Property (Word)

Returns or sets the amount of space (in points) to add below the contents of a single cell or all the cells in a table. Read/write  **Single** .


## Syntax

 _expression_ . **BottomPadding**

 _expression_ A variable that represents a **[Cell](cell-object-word.md)** object.


## Remarks

The setting of the  **BottomPadding** property for a single cell overrides the setting of the **BottomPadding** property for the entire table.


## Example

This example sets the bottom padding for the first table in the active document to 40 pixels.


```vb
ActiveDocument.Tables(1).BottomPadding = _ 
 PixelsToPoints(40, True)
```


## See also


#### Concepts


[Cell Object](cell-object-word.md)

