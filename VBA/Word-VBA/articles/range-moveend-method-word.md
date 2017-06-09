---
title: Range.MoveEnd Method (Word)
keywords: vbawd10.chm157155439
f1_keywords:
- vbawd10.chm157155439
ms.prod: word
api_name:
- Word.Range.MoveEnd
ms.assetid: 44aa26e6-7bb1-af51-8d23-244444e0795c
ms.date: 06/08/2017
---


# Range.MoveEnd Method (Word)

Moves the ending character position of a range. .


## Syntax

 _expression_ . **MoveEnd**( **_Unit_** , **_Count_** )

 _expression_ Required. A variable that represents a **[Range](range-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Unit_|Optional| **WdUnits**|The unit by which to move the ending character position.|
| _Count_|Optional| **Variant**|The number of units to move. If this number is positive, the ending character position is moved forward in the document. If this number is negative, the end is moved backward. If the ending position overtakes the starting position, the range collapses and both character positions move together. The default value is 1.|

## Remarks

This method returns an integer that indicates the number of units the range actually moved, or it returns 0 (zero) if the move was unsuccessful.


## Example

This example sets  _myRange_ to be equal to the second word in the active document. The **MoveEnd** method is used to move the ending position of _myRange_ (a range object) forward one word. After this macro is run, the second and third words in the document are selected.


```vb
If ActiveDocument.Words.Count >= 3 Then 
 Set myRange = ActiveDocument.Words(2) 
 With myRange 
 .MoveEnd Unit:=wdWord, Count:=1 
 .Select 
 End With 
End If
```


## See also


#### Concepts


[Range Object](range-object-word.md)

