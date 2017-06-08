---
title: Range.MoveUntil Method (Word)
keywords: vbawd10.chm157155443
f1_keywords:
- vbawd10.chm157155443
ms.prod: word
api_name:
- Word.Range.MoveUntil
ms.assetid: f0f44ae5-1d61-9e05-4095-a28091feda6f
ms.date: 06/08/2017
---


# Range.MoveUntil Method (Word)

Moves the specified range until one of the specified characters is found in the document.


## Syntax

 _expression_ . **MoveUntil**( **_Cset_** , **_Count_** )

 _expression_ Required. A variable that represents a **[Range](range-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cset_|Required| **Variant**|One or more characters. If any character in Cset is found before the Count value expires, the specified range is positioned as an insertion point immediately before that character. This argument is case sensitive.|
| _Count_|Optional| **Variant**|The maximum number of characters by which the specified range is to be moved. Can be a number or either the  **wdForward** or **wdBackward** constant. If Count is a positive number, the range is moved forward in the document, beginning at the end position. If it is a negative number, the range is moved backward, beginning at the start position. The default value is **wdForward** .|

## Remarks

This method returns the number of characters by which the specified range was moved, as a  **Long** value. If Count is greater than 0 (zero), this method returns the number of characters moved plus one. If Count is less than 0 (zero), this method returns the number of characters moved minus one. If no Cset characters are found, the range isn't not changed and the method returns 0 (zero).


## Example

This example moves  _myRange_ forward through the next 100 characters in the document until the character "t" is found.


```vb
Set myRange = ActiveDocument.Words(1) 
myRange.MoveUntil Cset:="t", Count:=100
```


## See also


#### Concepts


[Range Object](range-object-word.md)

