---
title: Selection.InsertFormula Method (Word)
keywords: vbawd10.chm158663186
f1_keywords:
- vbawd10.chm158663186
ms.prod: word
api_name:
- Word.Selection.InsertFormula
ms.assetid: a193c4ee-a667-04af-e22c-3a5b5bbc5c3b
ms.date: 06/08/2017
---


# Selection.InsertFormula Method (Word)

Inserts an = (Formula) field that contains a formula at the selection.


## Syntax

 _expression_ . **Formula**( **_Formula_** , **_NumberFormat_** )

 _expression_ Required. A variable that represents a **[Selection](selection-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Formula_|Optional| **Variant**|The mathematical formula you want the = (Formula) field to evaluate. Spreadsheet-type references to table cells are valid. For example, "=SUM(A4:C4)" specifies the first three values in the fourth row. For more information about the = (Formula) field, see Field codes:= (Formula) field.|
| _NumberFormat_|Optional| **Variant**|A format for the result of the = (Formula) field. For information about the types of formats you can apply, see Numeric Picture (\#) field switch.|

## Remarks

The formula replaces the selection, if the selection is not collapsed.

If you are using a spreadsheet application, such as Microsoft Office Excel, embedding all or part of a worksheet in a document is often easier than using the = (Formula) field in a table.

The Formula argument is optional only if the selection is in a cell and there is at least one cell that contains a value above or to the left of the cell that contains the insertion point. If the cells above the insertion point contain values, the inserted field is {=SUM(ABOVE)}; if the cells to the left of the insertion point contain values, the inserted field is {=SUM(LEFT)}. If both the cells above the insertion point and the cells to the left of it contain values, Microsoft Word uses the following rules to determine which SUM function to insert:


- If the cell immediately above the insertion point contains a value, Word inserts {=SUM(ABOVE)}.
    
- If the cell immediately above the insertion point does not contain a value but the cell immediately to the left of the insertion point does, Word inserts {=SUM(LEFT)}.
    
- If neither cell immediately above the insertion point nor the cell immediately below it contains a value, Word inserts {=SUM(ABOVE)}.
    
- If you don't specify  **Formula** and all the cells above and to the left of the insertion point are empty, using the = (Formula) field causes an error.
    

## Example

This example creates a table with three rows and three columns at the beginning of the active document and then calculates the average of all the numbers in the first column.


```vb
Set MyRange = ActiveDocument.Range(0, 0) 
Set myTable = ActiveDocument.Tables.Add(MyRange, 3, 3) 
With myTable 
 .Cell(1, 1).Range.InsertAfter "100" 
 .Cell(2, 1).Range.InsertAfter "50" 
 .Cell(3, 1).Select 
End With 
Selection.InsertFormula Formula:="=Average(Above)"
```

The example inserts a formula field that is subtracted from a value represented by the bookmark named "GrossSales." The result is formatted with a dollar sign.




```
Selection.Collapse Direction:=wdCollapseStart 
Selection.InsertFormula Formula:= "=GrossSales-45,000.00", _ 
 NumberFormat:="$#,##0.00"
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

