---
title: Range.MoveEndUntil Method (Word)
keywords: vbawd10.chm157155445
f1_keywords:
- vbawd10.chm157155445
ms.prod: word
api_name:
- Word.Range.MoveEndUntil
ms.assetid: 62ac37a2-1116-73de-dcd8-0ff74ae7803b
ms.date: 06/08/2017
---


# Range.MoveEndUntil Method (Word)

Moves the end position of the specified range until any of the specified characters are found in the document. If the movement is forward in the document, the range is expanded.


## Syntax

 _expression_ . **MoveEndUntil**( **_Cset_** , **_Count_** )

 _expression_ Required. A variable that represents a **[Range](range-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cset_|Required| **Variant**|One or more characters. This argument is case sensitive.|
| _Count_|Optional| **Variant**|The maximum number of characters by which the specified range is to be moved. Can be a number or either the  **wdForward** or **wdBackward** constant. If Count is a positive number, the range is moved forward in the document. If it is a negative number, the range is moved backward. The default value is **wdForward** .|

## Remarks

This method returns the number of characters by which the end position of the specified range was moved, as a  **Long** value. If Count is greater than 0 (zero), this method returns the number of characters moved plus 1. If Count is less than 0 (zero), this method returns the number of characters moved minus 1. If no Cset characters are found, the range isn't changed and the method returns 0 (zero). If the end position is moved backward to a point that precedes the original start position, the start position is set to the new ending position.


## Example

This example extends the selected text forward in the document until the letter "a" is found. The example then expands the selection by one character to include the letter "a".


```vb
With Selection.Range 
 .MoveEndUntil Cset:="a", Count:=wdForward 
 .MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdExtend 
End With
```


## See also


#### Concepts


[Range Object](range-object-word.md)

