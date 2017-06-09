---
title: Table.RightPadding Property (Word)
keywords: vbawd10.chm156303476
f1_keywords:
- vbawd10.chm156303476
ms.prod: word
api_name:
- Word.Table.RightPadding
ms.assetid: a41681da-9a11-9b45-fcff-495208a3ab25
ms.date: 06/08/2017
---


# Table.RightPadding Property (Word)

Returns or sets the amount of space (in points) to add to the right of the contents of all the cells in a table. Read/write  **Single** .


## Syntax

 _expression_ . **RightPadding**

 _expression_ Required. A variable that represents a **[Table](table-object-word.md)** object.


## Remarks

The setting of the  **RightPadding** property for a single cell overrides the setting of the **RightPadding** property for the entire table.


## Example

This example sets the right padding for the first table in the active document to 40 pixels.


```vb
ActiveDocument.Tables(1).RightPadding = _ 
 PixelsToPoints(40, False)
```


## See also


#### Concepts


[Table Object](table-object-word.md)

