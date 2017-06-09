---
title: Selection.MoveUntil Method (Word)
keywords: vbawd10.chm158662771
f1_keywords:
- vbawd10.chm158662771
ms.prod: word
api_name:
- Word.Selection.MoveUntil
ms.assetid: 888655d0-44ec-b589-bd0b-b3e193e413ef
ms.date: 06/08/2017
---


# Selection.MoveUntil Method (Word)

Moves the specified selection until one of the specified characters is found in the document.


## Syntax

 _expression_ . **MoveUntil**( **_Cset_** , **_Count_** )

 _expression_ Required. A variable that represents a **[Selection](selection-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cset_|Required| **Variant**|One or more characters. If any character in Cset is found before the Count value expires, the specified selection is positioned as an insertion point immediately before that character. This argument is case sensitive.|
| _Count_|Optional| **Variant**|The maximum number of characters by which the specified selection is to be moved. Can be a number or either the  **wdForward** or **wdBackward** constant. If Count is a positive number, the selection is moved forward in the document, beginning at the end position. If it is a negative number, the selection is moved backward, beginning at the start position. The default value is **wdForward** .|

## Remarks

This method returns the number of characters by which the specified selection was moved, as a  **Long** value. If Count is greater than 0 (zero), this method returns the number of characters moved plus one. If Count is less than 0 (zero), this method returns the number of characters moved minus one. If no Cset characters are found, the selection isn't not changed and the method returns 0 (zero).


## Example

This example moves the selection forward to the end of the active paragraph and then displays the number of characters by which the selection was moved.


```
x = Selection.MoveUntil(Cset:=Chr$(13), Count:=wdForward) 
MsgBox x-1 &; " character positions were moved"
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

