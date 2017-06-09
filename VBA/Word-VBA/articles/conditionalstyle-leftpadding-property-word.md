---
title: ConditionalStyle.LeftPadding Property (Word)
keywords: vbawd10.chm91029509
f1_keywords:
- vbawd10.chm91029509
ms.prod: word
api_name:
- Word.ConditionalStyle.LeftPadding
ms.assetid: 5bb8fdb1-a971-13bc-4977-b0ffdcb95116
ms.date: 06/08/2017
---


# ConditionalStyle.LeftPadding Property (Word)

Returns or sets the amount of space (in points) to add to the left of the contents of a single cell or all the cells in a table. Read/write  **Single** .


## Syntax

 _expression_ . **LeftPadding**

 _expression_ Required. A variable that represents a **[ConditionalStyle](conditionalstyle-object-word.md)** object.


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


[ConditionalStyle Object](conditionalstyle-object-word.md)

