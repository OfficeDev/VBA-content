---
title: Range.Offset Property (Excel)
keywords: vbaxl10.chm144169
f1_keywords:
- vbaxl10.chm144169
ms.prod: excel
api_name:
- Excel.Range.Offset
ms.assetid: dfbbd1a2-2f73-fd6a-6277-4584823f55a4
ms.date: 06/08/2017
---


# Range.Offset Property (Excel)

Returns a  **[Range](range-object-excel.md)** object that represents a range that?s offset from the specified range.


## Syntax

 _expression_ . **Offset**( **_RowOffset_** , **_ColumnOffset_** )

 _expression_ A variable that represents a **Range** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _RowOffset_|Optional| **Variant**|The number of rows (positive, negative, or 0 (zero)) by which the range is to be offset. Positive values are offset downward, and negative values are offset upward. The default value is 0.|
| _ColumnOffset_|Optional| **Variant**|The number of columns (positive, negative, or 0 (zero)) by which the range is to be offset. Positive values are offset to the right, and negative values are offset to the left. The default value is 0.|

## Example

This example activates the cell three columns to the right of and three rows down from the active cell on Sheet1.


```vb
Worksheets("Sheet1").Activate 
ActiveCell.Offset(rowOffset:=3, columnOffset:=3).Activate
```

This example assumes that Sheet1 contains a table that has a header row. The example selects the table, without selecting the header row. The active cell must be somewhere in the table before the example is run.




```vb
Set tbl = ActiveCell.CurrentRegion 
tbl.Offset(1, 0).Resize(tbl.Rows.Count - 1, _ 
 tbl.Columns.Count).Select 

```


## See also


#### Concepts


[Range Object](range-object-excel.md)

