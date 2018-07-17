---
title: Range.MoveStart Method (Word)
keywords: vbawd10.chm157155438
f1_keywords:
- vbawd10.chm157155438
ms.prod: word
api_name:
- Word.Range.MoveStart
ms.assetid: 9097c636-594d-8a2e-8209-dc0db850812a
ms.date: 06/08/2017
---


# Range.MoveStart Method (Word)

Moves the start position of the specified range.


## Syntax

 _expression_ . **MoveStart**( **_Unit_** , **_Count_** )

 _expression_ Required. A variable that represents a **[Range](range-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Unit_|Optional| **WdUnits**|The unit by which start position of the specified range is to be moved.|
| _Count_|Optional| **Variant**|The maximum number of units by which the specified range is to be moved. If Count is a positive number, the start position of the range is moved forward in the document. If it is a negative number, the start position is moved backward. If the start position is moved forward to a position beyond the end position, the range is collapsed and both the start and end positions are moved together. The default value is 1.|

### Return Value

Integer


## Remarks

This method returns an integer that indicates the number of units by which the start position or the range actually moved, or it returns 0 (zero) if the move was unsuccessful.


## Example

This example sets  _myRange_ to be equal to the second word in the active document. The example uses the **MoveStart** method to move the start position of _myRange_ (a **Range** object) backward one word. After this macro is run, the first and second words in the document are selected.


```vb
If ActiveDocument.Words.Count >= 2 Then 
 Set myRange = ActiveDocument.Words(2) 
 With myRange 
 .MoveStart Unit:=wdWord, Count:=-1 
 .Select 
 End With 
End If
```


## See also


#### Concepts


[Range Object](range-object-word.md)

