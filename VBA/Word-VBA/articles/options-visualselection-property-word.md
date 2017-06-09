---
title: Options.VisualSelection Property (Word)
keywords: vbawd10.chm162988436
f1_keywords:
- vbawd10.chm162988436
ms.prod: word
api_name:
- Word.Options.VisualSelection
ms.assetid: d3947a4c-0495-6211-7646-3b202855d35a
ms.date: 06/08/2017
---


# Options.VisualSelection Property (Word)

Returns or sets the selection behavior based on visual cursor movement in a right-to-left language document. Read/write  **WdVisualSelection** .


## Syntax

 _expression_ . **VisualSelection**

 _expression_ Required. A variable that represents an **[Options](options-object-word.md)** collection.


## Remarks

The  **CursorMovement** property must be set to **wdCursorMovementVisual** to use this property.


## Example

This example sets the selection behavior so that the selection wraps from line to line.


```vb
If Options.CursorMovement = wdCursorMovementVisual Then _ 
 Options.VisualSelection = wdVisualSelectionContinuous
```


## See also


#### Concepts


[Options Object](options-object-word.md)

