---
title: Range.UnMerge Method (Excel)
keywords: vbaxl10.chm144159
f1_keywords:
- vbaxl10.chm144159
ms.prod: excel
api_name:
- Excel.Range.UnMerge
ms.assetid: dfc49876-29b0-0b61-fe18-3953438f7452
ms.date: 06/08/2017
---


# Range.UnMerge Method (Excel)

Separates a merged area into individual cells.


## Syntax

 _expression_ . **UnMerge**

 _expression_ A variable that represents a **Range** object.


## Example

This example separates the merged range that contains cell A3.


```vb
With Range("a3") 
 If .MergeCells Then 
 .MergeArea.UnMerge 
 Else 
 MsgBox "not merged" 
 End If 
End With
```


## See also


#### Concepts


[Range Object](range-object-excel.md)

