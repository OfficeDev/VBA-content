---
title: Range.RowDifferences Method (Excel)
keywords: vbaxl10.chm144189
f1_keywords:
- vbaxl10.chm144189
ms.prod: excel
api_name:
- Excel.Range.RowDifferences
ms.assetid: 89030ca3-9f59-7426-d050-89dcabf00887
ms.date: 06/08/2017
---


# Range.RowDifferences Method (Excel)

Returns a  **[Range](range-object-excel.md)** object that represents all the cells whose contents are different from those of the comparison cell in each row.


## Syntax

 _expression_ . **RowDifferences**( **_Comparison_** )

 _expression_ A variable that represents a **Range** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Comparison_|Required| **Variant**|A single cell to compare with the specified range.|

### Return Value

Range


## Example

This example selects the cells in row one on Sheet1 whose contents are different from those of cell D1.


```vb
Worksheets("Sheet1").Activate 
Set c1 = ActiveSheet.Rows(1).RowDifferences( _ 
 comparison:=ActiveSheet.Range("D1")) 
c1.Select
```


## See also


#### Concepts


[Range Object](range-object-excel.md)

