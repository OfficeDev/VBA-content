---
title: Range.IsEqual Method (Word)
keywords: vbawd10.chm157155499
f1_keywords:
- vbawd10.chm157155499
ms.prod: word
api_name:
- Word.Range.IsEqual
ms.assetid: cd6269d9-4693-897d-d9b2-69f45c815ba3
ms.date: 06/08/2017
---


# Range.IsEqual Method (Word)

 **True** if the range to which this method is applied is equal to the range specified by the Range argument.


## Syntax

 _expression_ . **IsEqual**( **_Range_** )

 _expression_ Required. A variable that represents a **[Range](range-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Range_|Required| **Range**|The range to compare with the  **Range** object defined by expression.|

### Return Value

Boolean


## Remarks

This method compares the starting and ending character positions and the story type. If all three of these items are the same for both objects, the objects are equal.


## Example

This example compares  _Range1_ with _Range2_ to determine whether they're equal. If the two ranges are equal, the content of _Range1_ is deleted.


```vb
Set Range1 = Selection.Words(1) 
Set Range2 = ActiveDocument.Words(3) 
If Range1.IsEqual(Range:=Range2) = True Then 
 Range1.Delete 
End If
```


## See also


#### Concepts


[Range Object](range-object-word.md)

