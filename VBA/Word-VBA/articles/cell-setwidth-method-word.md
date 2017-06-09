---
title: Cell.SetWidth Method (Word)
keywords: vbawd10.chm156106954
f1_keywords:
- vbawd10.chm156106954
ms.prod: word
api_name:
- Word.Cell.SetWidth
ms.assetid: fd9fbeb1-a8c7-a6bf-1c9e-b63954848baf
ms.date: 06/08/2017
---


# Cell.SetWidth Method (Word)

Sets the width of columns or cells in a table.


## Syntax

 _expression_ . **SetWidth**( **_ColumnWidth_** , **_RulerStyle_** )

 _expression_ Required. A variable that represents a **[Cell](cell-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ColumnWidth_|Required| **Single**|The width of the specified column or columns, in points.|
| _RulerStyle_|Required| **WdRulerStyle**|Controls the way Word adjusts cell widths.|

## Remarks

The  **WdRulerStyle** behavior described above applies to left-aligned tables. The **WdRulerStyle** behavior for center- and right-aligned tables can be unexpected; in these cases, the **SetWidth** method should be used with care.


## Example

This example creates a table in a new document and sets the width of the first cell in the second row to 1.5 inches. The example preserves the widths of the other cells in the table.


```vb
Set newDoc = Documents.Add 
Set myTable = _ 
 newDoc.Tables.Add(Range:=Selection.Range, NumRows:=3, _ 
 NumColumns:=3) 
myTable.Cell(2,1).SetWidth _ 
 ColumnWidth:=InchesToPoints(1.5), _ 
 RulerStyle:=wdAdjustNone
```

This example sets the width of the cell that contains the insertion point to 36 points. The example also narrows the first column to preserve the position of the right edge of the table.




```vb
If Selection.Information(wdWithInTable) = True Then 
 Selection.Cells(1).SetWidth ColumnWidth:=36, _ 
 RulerStyle:=wdAdjustFirstColumn 
Else 
 MsgBox "The insertion point is not in a table." 
End If
```


## See also


#### Concepts


[Cell Object](cell-object-word.md)

