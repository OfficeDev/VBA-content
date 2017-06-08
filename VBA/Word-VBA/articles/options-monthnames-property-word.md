---
title: Options.MonthNames Property (Word)
keywords: vbawd10.chm162988434
f1_keywords:
- vbawd10.chm162988434
ms.prod: word
api_name:
- Word.Options.MonthNames
ms.assetid: 265bee60-26ac-a6f5-4950-494ce6eff215
ms.date: 06/08/2017
---


# Options.MonthNames Property (Word)

Returns or sets the direction for conversion between Hangul and Hanja. Read/write  **WdMonthNames** .


## Syntax

 _expression_ . **MonthNames**

 _expression_ Required. A variable that represents an **[Options](options-object-word.md)** collection.


## Example

This example sets Microsoft Word to convert from Hangul to Hanja by default.


```
Options.MultipleWordConversionsMode = wdHangulToHanja
```


## See also


#### Concepts


[Options Object](options-object-word.md)

