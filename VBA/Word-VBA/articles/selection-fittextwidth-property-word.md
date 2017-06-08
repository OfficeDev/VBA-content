---
title: Selection.FitTextWidth Property (Word)
keywords: vbawd10.chm158663664
f1_keywords:
- vbawd10.chm158663664
ms.prod: word
api_name:
- Word.Selection.FitTextWidth
ms.assetid: 7f7409b4-c533-9c21-2663-e4016416efb7
ms.date: 06/08/2017
---


# Selection.FitTextWidth Property (Word)

Returns or sets the width (in the current measurement units) in which Microsoft Word fits the text in the current selection. Read/write  **Single** .


## Syntax

 _expression_ . **FitTextWidth**

 _expression_ A variable that represents a **[Selection](selection-object-word.md)** object.


## Example

This example fits the current selection into a space five centimeters wide.


```
Selection.FitTextWidth = CentimetersToPoints(5)
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

