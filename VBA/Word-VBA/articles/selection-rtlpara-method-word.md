---
title: Selection.RtlPara Method (Word)
keywords: vbawd10.chm158663261
f1_keywords:
- vbawd10.chm158663261
ms.prod: word
api_name:
- Word.Selection.RtlPara
ms.assetid: b417897d-de70-6c3a-12cd-8786e12bdb43
ms.date: 06/08/2017
---


# Selection.RtlPara Method (Word)

Sets the reading order and alignment of the specified paragraphs to right-to-left.


## Syntax

 _expression_ . **RtlPara**

 _expression_ Required. A variable that represents a **[Selection](selection-object-word.md)** object.


## Remarks

For all selected paragraphs, this method sets the reading order to right-to-left. If a paragraph with a left-to-right reading order is also left-aligned, this method reverses its reading order and sets its paragraph alignment to right-aligned.

Use the  **ReadingOrder** property to change the reading order without affecting paragraph alignment.


## Example

This example sets the reading order and alignment of the current selection to right-to-left if the selection isn't styled as "Normal."


```vb
If Selection.Style <> "Normal" Then _ 
 Selection.RtlPara
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

