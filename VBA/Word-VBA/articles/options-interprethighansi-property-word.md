---
title: Options.InterpretHighAnsi Property (Word)
keywords: vbawd10.chm162988450
f1_keywords:
- vbawd10.chm162988450
ms.prod: word
api_name:
- Word.Options.InterpretHighAnsi
ms.assetid: c093469b-c9ef-0b37-fc40-7b1ae17ce72e
ms.date: 06/08/2017
---


# Options.InterpretHighAnsi Property (Word)

Returns or sets the high-ANSI text interpretation behavior. Read/write  **WdHighAnsiText** .


## Syntax

 _expression_ . **InterpretHighAnsi**

 _expression_ Required. A variable that represents an **[Options](options-object-word.md)** collection.


## Example

This example sets Word to interpret all high-ANSI text as East Asian characters.


```
Options.InterpretHighAnsi = wdHighAnsiIsFarEast
```


## See also


#### Concepts


[Options Object](options-object-word.md)

