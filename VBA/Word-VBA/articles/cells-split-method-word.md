---
title: Cells.Split Method (Word)
keywords: vbawd10.chm155844813
f1_keywords:
- vbawd10.chm155844813
ms.prod: word
api_name:
- Word.Cells.Split
ms.assetid: ed0b2594-a328-20d9-b352-5a59b8ef9d3a
ms.date: 06/08/2017
---


# Cells.Split Method (Word)

Splits a range of table cells.


## Syntax

 _expression_ . **Split**( **_NumRows_** , **_NumColumns_** , **_MergeBeforeSplit_** )

 _expression_ Required. A variable that represents a **[Cells](cells-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _NumRows_|Optional| **Variant**|The number of rows that the cell or group of cells is to be split into.|
| _NumColumns_|Optional| **Variant**|The number of columns that the cell or group of cells is to be split into.|
| _MergeBeforeSplit_|Optional| **Variant**| **True** to merge the cells with one another before splitting them.|

## Example

This example merges the selected cells into a single cell and then splits the cell into three cells in the same row.


```vb
If Selection.Information(wdWithInTable) = True Then 
 Selection.Cells.Split NumRows:=1, NumColumns:=3, _ 
 MergeBeforeSplit:= True 
End If
```


## See also


#### Concepts


[Cells Collection Object](cells-object-word.md)

