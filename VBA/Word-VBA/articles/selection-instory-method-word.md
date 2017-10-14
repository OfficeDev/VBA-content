---
title: Selection.InStory Method (Word)
keywords: vbawd10.chm158662781
f1_keywords:
- vbawd10.chm158662781
ms.prod: word
api_name:
- Word.Selection.InStory
ms.assetid: 29dae109-4361-f1ee-eb71-76f57ae186a3
ms.date: 06/08/2017
---


# Selection.InStory Method (Word)

 **True** if the selection to which this method is applied is in the same story as the range specified by the Range argument.


## Syntax

 _expression_ . **InStory**( **_Range_** )

 _expression_ Required. A variable that represents a **[Selection](selection-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Range_|Required| **Range**|The range whose story is compared with the story of the current selection.|

### Return Value

Boolean


## Remarks

A range can belong to only one story.


## Example

This example determines whether the selection is in the same story as the first paragraph in the active document. The message box displays "False" because the selection is in the primary header story and the first paragraph is in the main text story.


```vb
With ActiveDocument.ActiveWindow.View 
 .Type = wdPrintView 
 .SeekView = wdSeekCurrentPageHeader 
End With 
same = Selection.InStory(ActiveDocument.Paragraphs(1).Range) 
MsgBox same
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

