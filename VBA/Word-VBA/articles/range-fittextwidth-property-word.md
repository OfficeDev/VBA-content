---
title: Range.FitTextWidth Property (Word)
keywords: vbawd10.chm157155592
f1_keywords:
- vbawd10.chm157155592
ms.prod: word
api_name:
- Word.Range.FitTextWidth
ms.assetid: 6322c657-21db-bc45-e2d6-cb559edfc047
ms.date: 06/08/2017
---


# Range.FitTextWidth Property (Word)

Returns or sets the width (in the current measurement units) in which Microsoft Word fits the text in the current selection or range. Read/write  **Single** .


## Syntax

 _expression_ . **FitTextWidth**

 _expression_ Required. A variable that represents a **[Range](range-object-word.md)** object.


## Example

This example fits the current selection into a space five centimeters wide.


```
Selection.FitTextWidth = CentimetersToPoints(5)
```


## See also


#### Concepts


[Range Object](range-object-word.md)

