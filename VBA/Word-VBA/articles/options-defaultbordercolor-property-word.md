---
title: Options.DefaultBorderColor Property (Word)
keywords: vbawd10.chm162988376
f1_keywords:
- vbawd10.chm162988376
ms.prod: word
api_name:
- Word.Options.DefaultBorderColor
ms.assetid: 382f9780-d10d-925b-206d-d7c624b6b744
ms.date: 06/08/2017
---


# Options.DefaultBorderColor Property (Word)

Returns or sets the default 24-bit color to use for new  **[Border](border-object-word.md)** objects. Read/write.


## Syntax

 _expression_ . **DefaultBorderColor**

 _expression_ Required. A variable that represents an **[Options](options-object-word.md)** collection.


## Remarks

This property can be any valid  **WdColor** constant or a value returned by Visual Basic's **RGB** function.


## Example

This example sets the default color for new borders to teal.


```
Options.DefaultBorderColor = wdColorTeal
```


## See also


#### Concepts


[Options Object](options-object-word.md)

