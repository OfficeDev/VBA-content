---
title: Range.HorizontalInVertical Property (Word)
keywords: vbawd10.chm157155593
f1_keywords:
- vbawd10.chm157155593
ms.prod: word
api_name:
- Word.Range.HorizontalInVertical
ms.assetid: 1d0ec26c-62a1-26ef-1fef-f2ab497244cb
ms.date: 06/08/2017
---


# Range.HorizontalInVertical Property (Word)

Returns or sets the formatting for horizontal text set within vertical text. Read/write  **WdHorizontalInVerticalType** .


## Syntax

 _expression_ . **HorizontalInVertical**

 _expression_ Required. A variable that represents a **[Range](range-object-word.md)** object.


## Example

This example formats the current selection as horizontal text within a run of vertical text, fitting the text to the line width of the vertical text.


```
Selection.Range.HorizontalInVertical = wdHorizontalInVerticalFitInLine
```


## See also


#### Concepts


[Range Object](range-object-word.md)

