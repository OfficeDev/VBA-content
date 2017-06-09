---
title: Font.Position Property (Word)
keywords: vbawd10.chm156369039
f1_keywords:
- vbawd10.chm156369039
ms.prod: word
api_name:
- Word.Font.Position
ms.assetid: 34896092-bf63-3c9e-b18e-768e888feaeb
ms.date: 06/08/2017
---


# Font.Position Property (Word)

Returns or sets the position of text (in points) relative to the base line. Read/write  **Long** .


## Syntax

 _expression_ . **Position**

 _expression_ Required. A variable that represents a **[Font](font-object-word.md)** object.


## Remarks

A positive number raises the text, and a negative number lowers it.


## Example

This example lowers the selected text by 2 points.


```
Selection.Font.Position = -2
```


## See also


#### Concepts


[Font Object](font-object-word.md)

