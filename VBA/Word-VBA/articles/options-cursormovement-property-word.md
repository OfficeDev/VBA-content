---
title: Options.CursorMovement Property (Word)
keywords: vbawd10.chm162988435
f1_keywords:
- vbawd10.chm162988435
ms.prod: word
api_name:
- Word.Options.CursorMovement
ms.assetid: f73f8a6e-4a66-e3f8-7197-42d5c1f73bcf
ms.date: 06/08/2017
---


# Options.CursorMovement Property (Word)

Returns or sets how the insertion point progresses within bidirectional text. Read/write  **WdCursorMovement** .


## Syntax

 _expression_ . **CursorMovement**

 _expression_ Required. A variable that represents an **[Options](options-object-word.md)** collection.


## Example

This example sets the insertion point to progress to the next visually adjacent character as it moves through bidirectional text.


```
Options.CursorMovement = wdCursorMovementVisual
```


## See also


#### Concepts


[Options Object](options-object-word.md)

