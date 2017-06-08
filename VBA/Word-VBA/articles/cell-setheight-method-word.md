---
title: Cell.SetHeight Method (Word)
keywords: vbawd10.chm156106955
f1_keywords:
- vbawd10.chm156106955
ms.prod: word
api_name:
- Word.Cell.SetHeight
ms.assetid: 1c26425e-66f0-0558-5981-7161d730e8e1
ms.date: 06/08/2017
---


# Cell.SetHeight Method (Word)

Sets the height of table cells.


## Syntax

 _expression_ . **SetHeight**( **_RowHeight_** , **_HeightRule_** )

 _expression_ Required. A variable that represents a **[Cell](cell-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _RowHeight_|Required| **Variant**|The height of the row or rows, in points.|
| _HeightRule_|Required| **WdRowHeightRule**|The rule for determining the height of the specified cells.|

## Remarks

Setting the  **SetHeight** property of a **Cell** object automatically sets the property for the entire row.


## Example

This example sets the row height of the selected cells to at least 18 points.


```vb
If Selection.Information(wdWithInTable) = True Then 
 Selection.Cells.SetHeight RowHeight:=18, _ 
 HeightRule:=wdRowHeightAtLeast 
Else 
 MsgBox "The insertion point is not in a table." 
End If
```


## See also


#### Concepts


[Cell Object](cell-object-word.md)

