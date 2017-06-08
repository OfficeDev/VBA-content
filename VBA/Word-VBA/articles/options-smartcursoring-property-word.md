---
title: Options.SmartCursoring Property (Word)
keywords: vbawd10.chm162988491
f1_keywords:
- vbawd10.chm162988491
ms.prod: word
api_name:
- Word.Options.SmartCursoring
ms.assetid: 254a0a6d-ba83-3ca0-e7a7-38dea9b16436
ms.date: 06/08/2017
---


# Options.SmartCursoring Property (Word)

Returns or sets a  **Boolean** that represents whether smart cursoring is enabled. **True** enables smart cursoring.


## Syntax

 _expression_ . **SmartCursoring**

 _expression_ An expression that returns an **Options** object.


## Remarks

The  **SmartCursoring** property corresponds to the **Use Smart Cursoring** option in the **Edit** tab of the **Options** dialog box, which is selected by default.

When the  **SmartCursoring** property is **True** , scrolling in a document by using the PAGE DOWN key will move the cursor to the current page. If the **SmartCursoring** property is **False** , the cursor remains in the last edited position.


## Example

The following example disables smart cursoring.


```vb
Options.SmartCursoring = False
```


## See also


#### Concepts


[Options Object](options-object-word.md)

