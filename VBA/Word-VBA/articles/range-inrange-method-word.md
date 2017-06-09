---
title: Range.InRange Method (Word)
keywords: vbawd10.chm157155454
f1_keywords:
- vbawd10.chm157155454
ms.prod: word
api_name:
- Word.Range.InRange
ms.assetid: 8d6b2093-7720-b100-6e9e-6be761cabaf5
ms.date: 06/08/2017
---


# Range.InRange Method (Word)

Returns  **True** if the range to which the method is applied is contained in the range specified by the Range argument.


## Syntax

 _expression_ . **InRange**( **_Range_** )

 _expression_ Required. A variable that represents a **[Range](range-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Range_|Required| **Range**|Specifies the range that this method uses to determine if it is contained within the specified  **Range** object.|

### Return Value

Boolean


## Remarks

This method determines whether the range returned by expression is contained in the specified Range by comparing the starting and ending character positions and the story type.


## Example

This example sets  _myRange_ equal to the first word in the active document. If _myRange_ is not contained in the selected range, _myRange_ is selected.


```vb
Set myRange = ActiveDocument.Words(1) 
If myRange.InRange(Selection.Range) = False Then myRange.Select
```


## See also


#### Concepts


[Range Object](range-object-word.md)

