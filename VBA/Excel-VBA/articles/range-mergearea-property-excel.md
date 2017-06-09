---
title: Range.MergeArea Property (Excel)
keywords: vbaxl10.chm144160
f1_keywords:
- vbaxl10.chm144160
ms.prod: excel
api_name:
- Excel.Range.MergeArea
ms.assetid: 68586bba-fa9c-e0d4-0eae-a08613551a2c
ms.date: 06/08/2017
---


# Range.MergeArea Property (Excel)

Returns a  **Range** object that represents the merged range containing the specified cell. If the specified cell isn't in a merged range, this property returns the specified cell. Read-only **Variant** .


## Syntax

 _expression_ . **MergeArea**

 _expression_ A variable that represents a **Range** object.


## Remarks

The  **MergeArea** property only works on a single-cell range.


## Example

This example sets the value of the merged range that contains cell A3.


```vb
Set ma = Range("a3").MergeArea 
If ma.Address = "$A$3" Then 
 MsgBox "not merged" 
Else 
 ma.Cells(1, 1).Value = "42" 
End If
```


## See also


#### Concepts


[Range Object](range-object-excel.md)

