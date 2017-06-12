---
title: Selection.MoveWhile Method (Word)
keywords: vbawd10.chm158662768
f1_keywords:
- vbawd10.chm158662768
ms.prod: word
api_name:
- Word.Selection.MoveWhile
ms.assetid: ba35991c-2ae3-e78f-7538-c102149cf392
ms.date: 06/08/2017
---


# Selection.MoveWhile Method (Word)

Moves the specified selection while any of the specified characters are found in the document.


## Syntax

 _expression_ . **MoveWhile**( **_Cset_** , **_Count_** )

 _expression_ Required. A variable that represents a **[Selection](selection-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cset_|Required| **Variant**|One or more characters. This argument is case sensitive.|
| _Count_|Optional| **Variant**|The maximum number of characters by which the specified selection is to be moved. Can be a number or either the  **wdForward** or **wdBackward** constant. If Count is a positive number, the specified selection is moved forward in the document, beginning at the end position. If it is a negative number, the selection is moved backward, beginning at the start position. The default value is **wdForward** .|

## Remarks

While any character in Cset is found, the specified selection is moved. The resulting  **Selection** object is positioned as an insertion point after whatever Cset characters were found. This method returns the number of characters by which the specified selection was moved, as a **Long** value. If no Cset characters are found, the selection isn't changed and the method returns 0 (zero).


## Example

This example moves the selection after consecutive tabs.


```
Selection.MoveWhile Cset:=vbTab, Count:=wdForward
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

