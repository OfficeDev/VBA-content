---
title: Range.Relocate Method (Word)
keywords: vbawd10.chm157155507
f1_keywords:
- vbawd10.chm157155507
ms.prod: word
api_name:
- Word.Range.Relocate
ms.assetid: 2df77535-627f-d8ba-6ea2-15676b24221c
ms.date: 06/08/2017
---


# Range.Relocate Method (Word)

In outline view, moves the paragraphs within the specified range after the next visible paragraph or before the previous visible paragraph.


## Syntax

 _expression_ . **Relocate**( **_Direction_** )

 _expression_ Required. A variable that represents a **[Range](range-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Direction_|Required| **WdRelocate**|The direction of the move.|

## Remarks

Body text moves with a heading only if the body text is collapsed in outline view or if it is part of the range.


## Example

This example moves the third, fourth, and fifth paragraphs in the active document below the next (sixth) paragraph.


```vb
theStart = ActiveDocument.Paragraphs(3).Range.Start 
theEnd = ActiveDocument.Paragraphs(5).Range.End 
Set myRange = ActiveDocument.Range(Start:=theStart, End:=theEnd) 
ActiveDocument.ActiveWindow.View.Type = wdOutlineView 
myRange.Relocate Direction:=wdRelocateDown
```

This example moves the first paragraph in the selection above the previous paragraph.




```vb
ActiveDocument.ActiveWindow.View.Type = wdOutlineView 
Selection.Paragraphs(1).Range.Relocate Direction:=wdRelocateUp
```


## See also


#### Concepts


[Range Object](range-object-word.md)

