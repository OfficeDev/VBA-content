---
title: Range.MoveStartUntil Method (Word)
keywords: vbawd10.chm157155444
f1_keywords:
- vbawd10.chm157155444
ms.prod: word
api_name:
- Word.Range.MoveStartUntil
ms.assetid: 2506e3ec-593c-27ba-69b0-230351094f64
ms.date: 06/08/2017
---


# Range.MoveStartUntil Method (Word)

Moves the start position of the specified range until one of the specified characters is found in the document.


## Syntax

 _expression_ . **MoveStartUntil**( **_Cset_** , **_Count_** )

 _expression_ Required. A variable that represents a **[Range](range-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cset_|Required| **Variant**|One or more characters. This argument is case sensitive.|
| _Count_|Optional| **Variant**|The maximum number of characters by which the specified range is to be moved. Can be a number or either the  **wdForward** or **wdBackward** constant. If Count is a positive number, the range is moved forward in the document. If it is a negative number, the range is moved backward. The default value is **wdForward** .|

## Remarks

If the movement is backward through the document, the range is expanded.

This method returns the number of characters by which the start position of the specified range moved, as a  **Long** value. If Count is greater than 0 (zero), this method returns the number of characters moved plus 1. If Count is less than 0 (zero), this method returns the number of characters moved minus 1. If no Cset characters are found, the specified range isn't changed and the method returns 0 (zero). If the start position is moved forward to a point beyond the end position, the range is collapsed and both the start and end positions are moved together.


## Example

If there is a dollar sign character ($) in the first paragraph in the selected text, this example moves  _myRange_ just before the dollar sign.


```vb
Set myRange = Selection.Paragraphs(1).Range 
leng = myRange.End - myRange.Start 
myRange.Collapse Direction:=wdCollapseStart 
myRange.MoveStartUntil Cset:="$", Count:=leng
```


## See also


#### Concepts


[Range Object](range-object-word.md)

