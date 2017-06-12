---
title: Cell.Formula Method (Word)
keywords: vbawd10.chm156106953
f1_keywords:
- vbawd10.chm156106953
ms.prod: word
api_name:
- Word.Cell.Formula
ms.assetid: 0fec018a-5a6f-f5ec-ed1c-a963e53c27b3
ms.date: 06/08/2017
---


# Cell.Formula Method (Word)

Inserts an = (Formula) field that contains the specified formula into a table cell.


## Syntax

 _expression_ . **Formula**( **_Formula_** , **_NumFormat_** )

 _expression_ Required. A variable that represents a **[Cell](cell-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Formula_|Optional| **Variant**|The mathematical formula you want the = (Formula) field to evaluate. Spreadsheet-type references to table cells are valid. For example, "=SUM(A4:C4)" specifies the first three values in the fourth row. For more information about the = (Formula) field, see Field codes:= (Formula) field.|
| _NumFormat_|Optional| **Variant**|A format for the result of the = (Formula) field. For information about the types of formats you can apply, see Numeric Picture (\#) field switch.|

## Remarks

 **Formula** is optional as long as there is at least one cell that contains a value above or to the left of the cell that contains the insertion point. If the cells above the insertion point contain values, the inserted field is {=SUM(ABOVE)}; if the cells to the left of the insertion point contain values, the inserted field is {=SUM(LEFT)}. If both the cells above the insertion point and the cells to the left of the insertion point contain values, Microsoft Word uses the following rules to determine which SUM function to insert:


- If the cell immediately above the insertion point contains a value, Word inserts {=SUM(ABOVE)}.
    
- If the cell immediately above the insertion point doesn't contain a value and the cell immediately to the left of it does, Word inserts {=SUM(LEFT)}.
    
- If neither adjoining cell contains a value, Word inserts {=SUM(ABOVE)}.
    
- If you don't specify  **Formula** and all the cells above and to the left of the insertion point are empty, the result of the field is an error.
    

## Example

This example creates a 3x3 table at the beginning of the active document and then averages the numbers in the first column.


```vb
Set myRange = ActiveDocument.Range(0, 0) 
Set myTable = ActiveDocument.Tables.Add(myRange, 3, 3) 
With myTable 
 .Cell(1,1).Range.InsertAfter "100" 
 .Cell(2,1).Range.InsertAfter "50" 
 .Cell(3,1).Formula Formula:="=Average(Above)" 
End With
```

This example inserts a formula at the insertion point that determines the largest number in the cells above the selected cell.




```vb
Selection.Collapse Direction:=wdCollapseStart 
If Selection.Information(wdWithInTable) = True Then 
 Selection.Cells(1).Formula Formula:="=Max(Above)" 
Else 
 MsgBox "The insertion point is not in a table." 
End If
```


## See also


#### Concepts


[Cell Object](cell-object-word.md)

