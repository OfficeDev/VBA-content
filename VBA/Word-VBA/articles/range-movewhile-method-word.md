---
title: Range.MoveWhile Method (Word)
keywords: vbawd10.chm157155440
f1_keywords:
- vbawd10.chm157155440
ms.prod: word
api_name:
- Word.Range.MoveWhile
ms.assetid: 282464eb-60e6-df03-344f-6e666af8b01f
ms.date: 06/08/2017
---


# Range.MoveWhile Method (Word)

Moves the specified range while any of the specified characters are found in the document.


## Syntax

 _expression_ . **MoveWhile**( **_Cset_** , **_Count_** )

 _expression_ Required. A variable that represents a **[Range](range-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cset_|Required| **Variant**|One or more characters. This argument is case sensitive.|
| _Count_|Optional| **Variant**|The maximum number of characters by which the specified range is to be moved. Can be a number or either the  **wdForward** or **wdBackward** constant. If Count is a positive number, the specified range is moved forward in the document, beginning at the end position. If it is a negative number, the range is moved backward, beginning at the start position. The default value is **wdForward** .|

## Remarks

While any character in Cset is found, the specified range is moved. The resulting  **Range** object is positioned as an insertion point after whatever Cset characters were found. This method returns the number of characters by which the specified range was moved, as a **Long** value. If no Cset characters are found, the range isn't changed and the method returns 0 (zero).


## Example

This example moves  _aRange_ while any of the following (uppercase or lowercase) letters are found: "a", "t", or "i".


```vb
Dim aRange As Range 
Set aRange = ActiveDocument.Characters(1) 
aRange.MoveWhile Cset:="atiATI", Count:=wdForward
```


## See also


#### Concepts


[Range Object](range-object-word.md)

