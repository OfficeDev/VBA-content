---
title: Range.MoveEndWhile Method (Word)
keywords: vbawd10.chm157155442
f1_keywords:
- vbawd10.chm157155442
ms.prod: word
api_name:
- Word.Range.MoveEndWhile
ms.assetid: 9fab0517-a66a-2ae3-1900-77047b61cafa
ms.date: 06/08/2017
---


# Range.MoveEndWhile Method (Word)

Moves the ending character position of a range while any of the specified characters are found in the document.


## Syntax

 _expression_ . **MoveEndWhile**( **_Cset_** , **_Count_** )

 _expression_ Required. A variable that represents a **[Range](range-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cset_|Required| **Variant**|One or more characters. This argument is case sensitive.|
| _Count_|Optional| **Variant**|The maximum number of characters by which the range is to be moved. Can be a number or either the  **wdForward** or **wdBackward** constant. If Count is a positive number, the range is moved forward in the document. If it is a negative number, the range is moved backward. The default value is **wdForward** .|

## Remarks

While any character in Cset is found, the end position of the specified range is moved. This method returns the number of characters that the end position of the range moved as a  **Long** value. If no Cset characters are found, the range isn't changed and the method returns 0 (zero). If the end position is moved backward to a point that precedes the original start position, the start position is set to the new end position.


## Example

This example moves the end position of the selected range forward while the space character is found.


```
Selection.Range.MoveEndWhile Cset:=" ", Count:=wdForward
```

This example moves the end position of the selected range forward while Count is less than or equal to 10 and any letter from "a" through "h" is found.




```
Selection.Range.MoveEndWhile Cset:="abcdefgh", Count:=10
```


## See also


#### Concepts


[Range Object](range-object-word.md)

