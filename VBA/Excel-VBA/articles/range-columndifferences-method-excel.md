---
title: Range.ColumnDifferences Method (Excel)
keywords: vbaxl10.chm144100
f1_keywords:
- vbaxl10.chm144100
ms.prod: excel
api_name:
- Excel.Range.ColumnDifferences
ms.assetid: 483995e1-9c8d-c171-4c72-17afd5918d49
ms.date: 06/08/2017
---


# Range.ColumnDifferences Method (Excel)

Returns a  **[Range](range-object-excel.md)** object that represents all the cells whose contents are different from the comparison cell in each column.


## Syntax

 _expression_ . **ColumnDifferences**( **_Comparison_** )

 _expression_ A variable that represents a **Range** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Comparison_|Required| **Variant**|A single cell to compare to the specified range.|

### Return Value

Range


## Example

This example selects the cells in column A on Sheet1 whose contents are different from cell A4.


```vb
Sub CompDiff() 
'Setting up data to be compared 
 Range("A1").Value = "Rod" 
 Range("A2").Value = "Bill" 
 Range("A3").Value = "John" 
 Range("A4").Value = "Rod" 
 Range("A5").Value = "Kelly" 
 Range("A6").Value = "Rod" 
 Range("A7").Value = "Paddy" 
 Range("A8").Value = "Rod" 
 Range("A9").Value = "Rod" 
 Range("A10").Value = "Rod" 
 
'Code to do the comparison, selects the values that are unlike A1 
Worksheets("Sheet1").Activate 
Set r1 = ActiveSheet.Columns("A").ColumnDifferences( _ 
 Comparison:=ActiveSheet.Range("A1")) 
r1.Select 
End Sub
```


## See also


#### Concepts


[Range Object](range-object-excel.md)

