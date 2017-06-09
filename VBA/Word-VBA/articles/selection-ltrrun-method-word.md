---
title: Selection.LtrRun Method (Word)
keywords: vbawd10.chm158663257
f1_keywords:
- vbawd10.chm158663257
ms.prod: word
api_name:
- Word.Selection.LtrRun
ms.assetid: e2b905f1-3ce1-ce51-bc9f-c5325fa0e9af
ms.date: 06/08/2017
---


# Selection.LtrRun Method (Word)

Sets the reading order and alignment of the specified run to left-to-right.


## Syntax

 _expression_ . **LtrRun**

 _expression_ Required. A variable that represents a **[Selection](selection-object-word.md)** object.


## Remarks

For the specified run, this method sets the reading order to left-to-right. If a paragraph in the run with a right-to-left reading order is also right-aligned, this method reverses its reading order and sets its paragraph alignment to left-aligned.

Use the  **ReadingOrder** property to change the reading order without affecting paragraph alignment.


## Example

This example sets the reading order and alignment of the specified run to left-to-right if the selection is styled as "Normal."


```vb
If Selection.Style = "Normal" Then _ 
 Selection.LtrRun
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

