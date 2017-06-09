---
title: TableStyle.LeftPadding Property (Word)
keywords: vbawd10.chm244776965
f1_keywords:
- vbawd10.chm244776965
ms.prod: word
api_name:
- Word.TableStyle.LeftPadding
ms.assetid: e6b02546-7418-3df1-0d96-b6ec7b52f49d
ms.date: 06/08/2017
---


# TableStyle.LeftPadding Property (Word)

Returns or sets the amount of space (in points) to add to the left of the contents of all the cells in a table. Read/write  **Single** .


## Syntax

 _expression_ . **LeftPadding**

 _expression_ Required. A variable that represents a **[TableStyle](tablestyle-object-word.md)** object.


## Remarks

The setting of the  **LeftPadding** property for a single cell overrides the setting of the **LeftPadding** property for the entire table.


## Example

This example sets the left padding for the first table in the active document to 40 pixels.


```vb
ActiveDocument.Tables(1).LeftPadding = _ 
 PixelsToPoints(40, False)
```


## See also


#### Concepts


[TableStyle Object](tablestyle-object-word.md)

