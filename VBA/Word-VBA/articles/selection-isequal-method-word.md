---
title: Selection.IsEqual Method (Word)
keywords: vbawd10.chm158662827
f1_keywords:
- vbawd10.chm158662827
ms.prod: word
api_name:
- Word.Selection.IsEqual
ms.assetid: 57ca55bc-17cf-054c-81dd-aa6d1e536cd8
ms.date: 06/08/2017
---


# Selection.IsEqual Method (Word)

 **True** if the selection to which this method is applied is equal to the range specified by the Range argument.


## Syntax

 _expression_ . **IsEqual**( **_Range_** )

 _expression_ Required. A variable that represents a **[Selection](selection-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Range_|Required| **Range**|The range to compare with the current selection.|

### Return Value

Boolean


## Remarks

This method compares the starting and ending character positions and the story type. If all three of these items are the same for both the selection and the object specified by the Range argument, the objects are equal.


## Example

This example compares the selection with the second paragraph in the active document. If the selection isn't equal to the second paragraph, the second paragraph is selected.


```vb
If Selection.IsEqual(ActiveDocument _ 
 .Paragraphs(2).Range) = False Then 
 ActiveDocument.Paragraphs(2).Range.Select 
End If
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

