---
title: Range.SpecialCells Method (Excel)
keywords: vbaxl10.chm144203
f1_keywords:
- vbaxl10.chm144203
ms.prod: excel
api_name:
- Excel.Range.SpecialCells
ms.assetid: 30c2035c-34e3-3b1a-f243-69a9fed97f3b
ms.date: 06/08/2017
---


# Range.SpecialCells Method (Excel)

Returns a  **[Range](range-object-excel.md)** object that represents all the cells that match the specified type and value.


## Syntax

 _expression_ . **SpecialCells**( **_Type_** , **_Value_** )

 _expression_ A variable that represents a **Range** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Type_|Required| **[XlCellType](xlcelltype-enumeration-excel.md)**|The cells to include.|
| _Value_|Optional| **Variant**|If  _Type_ is either **xlCellTypeConstants** or **xlCellTypeFormulas** , this argument is used to determine which types of cells to include in the result. These values can be added together to return more than one type. The default is to select all constants or formulas, no matter what the type.|

### Return Value

Range


## Remarks





|**XlCellType constants**|**Value**|
|:-----|:-----|
| **xlCellTypeAllFormatConditions** . Cells of any format|-4172|
| **xlCellTypeAllValidation** . Cells having validation criteria|-4174|
| **xlCellTypeBlanks** . Empty cells|4|
| **xlCellTypeComments** . Cells containing notes|-4144|
| **xlCellTypeConstants** . Cells containing constants|2|
| **xlCellTypeFormulas** . Cells containing formulas|-4123|
| **xlCellTypeLastCell** . The last cell in the used range|11|
| **xlCellTypeSameFormatConditions** . Cells having the same format|-4173|
| **xlCellTypeSameValidation** . Cells having the same validation criteria|-4175|
| **xlCellTypeVisible** . All visible cells|12|


|** XlSpecialCellsValue constants**|**Value**|
|:-----|:-----|
| **xlErrors**|16|
| **xlLogical**|4|
| **xlNumbers**|1|
| **xlTextValues**|2|

## Example

This example selects the last cell in the used range of Sheet1.


```vb
Worksheets("Sheet1").Activate 
ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Activate
```


## See also


#### Concepts


[Range Object](range-object-excel.md)

