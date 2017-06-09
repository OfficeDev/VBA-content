---
title: Range.Collapse Method (Word)
keywords: vbawd10.chm157155429
f1_keywords:
- vbawd10.chm157155429
ms.prod: word
api_name:
- Word.Range.Collapse
ms.assetid: fa5cae70-f047-e300-52f7-bd75d9c613da
ms.date: 06/08/2017
---


# Range.Collapse Method (Word)

Collapses a range or selection to the starting or ending position. After a range or selection is collapsed, the starting and ending points are equal.


## Syntax

 _expression_ . **Collapse**( **_Direction_** )

 _expression_ Required. A variable that represents a **[Range](range-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Direction_|Optional| **Variant**|The direction in which to collapse the range or selection. Can be either of the following  **WdCollapseDirection** constants: **wdCollapseEnd** or **wdCollapseStart** . The default value is **wdCollapseStart** .|

## Remarks

If you use  **wdCollapseEnd** to collapse a range that refers to an entire paragraph, the range is located after the ending paragraph mark (the beginning of the next paragraph). However, you can move the range back one character by using the **MoveEnd** method after the range is collapsed, as shown in the following example.


```vb
Set myRange = ActiveDocument.Paragraphs(1).Range 
myRange.Collapse Direction:=wdCollapseEnd 
myRange.MoveEnd Unit:=wdCharacter, Count:=-1
```


## Example

This example sets myRange equal to the contents of the active document, collapses myRange, and then inserts a 2x2 table at the end of the document.


```vb
Set myRange = ActiveDocument.Content 
myRange.Collapse Direction:=wdCollapseEnd 
ActiveDocument.Tables.Add Range:=myRange, NumRows:=2, NumColumns:=2
```


## See also


#### Concepts


[Range Object](range-object-word.md)

