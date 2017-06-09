---
title: Range.Resize Property (Excel)
keywords: vbaxl10.chm144187
f1_keywords:
- vbaxl10.chm144187
ms.prod: excel
api_name:
- Excel.Range.Resize
ms.assetid: 05af0539-8aa3-c83c-1972-dfac618929b9
ms.date: 06/08/2017
---


# Range.Resize Property (Excel)

Resizes the specified range. Returns a  **[Range](range-object-excel.md)** object that represents the resized range.


## Syntax

 _expression_ . **Resize**( **_RowSize_** , **_ColumnSize_** )

 _expression_ An expression that returns a **Range** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _RowSize_|Optional| **Variant**|The number of rows in the new range. If this argument is omitted, the number of rows in the range remains the same.|
| _ColumnSize_|Optional| **Variant**|The number of columns in the new range. If this argument is omitted, the number of columns in the range remains the same.|

### Return Value

Range


## Example

This example resizes the selection on Sheet1 to extend it by one row and one column.


```vb
Worksheets("Sheet1").Activate 
numRows = Selection.Rows.Count 
numColumns = Selection.Columns.Count 
Selection.Resize(numRows + 1, numColumns + 1).Select
```

This example assumes that you have a table on Sheet1 that has a header row. The example selects the table, without selecting the header row. The active cell must be somewhere in the table before you run the example.




```vb
Set tbl = ActiveCell.CurrentRegion 
tbl.Offset(1, 0).Resize(tbl.Rows.Count - 1, _ 
 tbl.Columns.Count).Select
```


## See also


#### Concepts


[Range Object](range-object-excel.md)

