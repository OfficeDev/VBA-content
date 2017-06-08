---
title: Selection.RtlRun Method (Word)
keywords: vbawd10.chm158663256
f1_keywords:
- vbawd10.chm158663256
ms.prod: word
api_name:
- Word.Selection.RtlRun
ms.assetid: 759a16cd-24d7-7c0a-6315-47d395560c73
ms.date: 06/08/2017
---


# Selection.RtlRun Method (Word)

Sets the reading order and alignment of the specified run to right-to-left.


## Syntax

 _expression_ . **RtlRun**

 _expression_ Required. A variable that represents a **[Selection](selection-object-word.md)** object.


## Remarks

For the specified run, this method sets the reading order to right-to-left. If a paragraph in the run with a left-to-right reading order is also left-aligned, this method reverses its reading order and sets its paragraph alignment to right-aligned.

Use the  **ReadingOrder** property to change the reading order without affecting paragraph alignment.


## Example

This example sets the reading order and alignment of the specified run to right-to-left if the selection isn't styled as "Normal."


```vb
If Selection.Style <> "Normal" Then _ 
 Selection.RtlRun
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

