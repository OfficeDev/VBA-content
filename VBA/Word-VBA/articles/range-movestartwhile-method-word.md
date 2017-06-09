---
title: Range.MoveStartWhile Method (Word)
keywords: vbawd10.chm157155441
f1_keywords:
- vbawd10.chm157155441
ms.prod: word
api_name:
- Word.Range.MoveStartWhile
ms.assetid: d0cff673-9248-88ae-7624-a838ce104e4b
ms.date: 06/08/2017
---


# Range.MoveStartWhile Method (Word)

Moves the start position of the specified range while any of the specified characters are found in the document.


## Syntax

 _expression_ . **MoveStartWhile**( **_Cset_** , **_Count_** )

 _expression_ Required. A variable that represents a **[Range](range-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cset_|Required| **Variant**|One or more characters. This argument is case sensitive.|
| _Count_|Optional| **Variant**|The maximum number of characters by which the specified range is to be moved. Can be a number or either the  **wdForward** or **wdBackward** constant. If Count is a positive number, the range is moved forward in the document. If it is a negative number, the range is moved backward. The default value is **wdForward** .|

## Remarks

While any character in Cset is found, the start position of the range is moved. This method returns the number of characters that the start position of the range moved as a  **Long** value. If not Cset characters are found, the range isn't changed and the method returns 0 (zero). If the start position is moved forward to a position beyond the original end position, the end position is set to the new start position.


## Example

This example moves the start position of the selected range backward through the document while the space character is found.


```
Selection.Range.MoveStartWhile Cset:=" ", Count:=wdBackward
```

This example moves the start position of the selected range backward through the document while Count is less than or equal to 10 and any letter from "a" through "h" is found.




```
Selection.Range.MoveStartWhile Cset:="abcdefgh", Count:=-10
```


## See also


#### Concepts


[Range Object](range-object-word.md)

