---
title: ConditionalStyle.RightPadding Property (Word)
keywords: vbawd10.chm91029510
f1_keywords:
- vbawd10.chm91029510
ms.prod: word
api_name:
- Word.ConditionalStyle.RightPadding
ms.assetid: ebdaeb98-9d4b-039f-0ef0-4e0c7a611f1e
ms.date: 06/08/2017
---


# ConditionalStyle.RightPadding Property (Word)

Returns or sets the amount of space (in points) to add to the right of the contents of a single cell or all the cells in a table. Read/write  **Single** .


## Syntax

 _expression_ . **RightPadding**

 _expression_ Required. A variable that represents a **[ConditionalStyle](conditionalstyle-object-word.md)** object.


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


[ConditionalStyle Object](conditionalstyle-object-word.md)

