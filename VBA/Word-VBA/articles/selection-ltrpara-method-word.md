---
title: Selection.LtrPara Method (Word)
keywords: vbawd10.chm158663262
f1_keywords:
- vbawd10.chm158663262
ms.prod: word
api_name:
- Word.Selection.LtrPara
ms.assetid: 992886b8-44e3-3b1f-cc6d-7b16e1c58aef
ms.date: 06/08/2017
---


# Selection.LtrPara Method (Word)

Sets the reading order and alignment of the specified paragraphs to left-to-right.


## Syntax

 _expression_ . **LtrPara**

 _expression_ Required. A variable that represents a **[Selection](selection-object-word.md)** object.


## Remarks

For all selected paragraphs, this method sets the reading order to left-to-right. If a paragraph with a right-to-left reading order is also right-aligned, this method reverses its reading order and sets its paragraph alignment to left-aligned.

Use the  **ReadingOrder** property to change the reading order without affecting paragraph alignment.


## Example

This example sets the reading order and alignment of the current selection to left-to-right if the selection is styled as "Normal."


```vb
If Selection.Style = "Normal" Then _ 
 Selection.LtrPara
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

