---
title: TabStop.Clear Method (Word)
keywords: vbawd10.chm156500168
f1_keywords:
- vbawd10.chm156500168
ms.prod: word
api_name:
- Word.TabStop.Clear
ms.assetid: 5337df07-97a5-2dfe-97b3-7277649b4701
ms.date: 06/08/2017
---


# TabStop.Clear Method (Word)

Removes the specified custom tab stop.


## Syntax

 _expression_ . **Clear**

 _expression_ Required. A variable that represents a **[TabStop](tabstop-object-word.md)** object.


## Example

This example clears the first custom tab in the first paragraph of the active document.


```vb
ActiveDocument.Paragraphs(1).TabStops(1).Clear
```


## See also


#### Concepts


[TabStop Object](tabstop-object-word.md)

