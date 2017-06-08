---
title: Options.MultipleWordConversionsMode Property (Word)
keywords: vbawd10.chm162988375
f1_keywords:
- vbawd10.chm162988375
ms.prod: word
api_name:
- Word.Options.MultipleWordConversionsMode
ms.assetid: 4200229d-9a37-4b51-6cdc-e24e241aceff
ms.date: 06/08/2017
---


# Options.MultipleWordConversionsMode Property (Word)

Returns or sets the direction for conversion between Hangul and Hanja. Read/write  **WdMultipleWordConversionsMode** .


## Syntax

 _expression_ . **MultipleWordConversionsMode**

 _expression_ Required. A variable that represents an **[Options](options-object-word.md)** collection.


## Example

This example sets Microsoft Word to convert from Hangul to Hanja by default.


```
Options.MultipleWordConversionsMode = wdHangulToHanja
```


## See also


#### Concepts


[Options Object](options-object-word.md)

