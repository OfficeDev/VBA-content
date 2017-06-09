---
title: PivotField.Subtotals Property (Excel)
keywords: vbaxl10.chm240094
f1_keywords:
- vbaxl10.chm240094
ms.prod: excel
api_name:
- Excel.PivotField.Subtotals
ms.assetid: 1086c36f-e792-b2a5-848a-efd2c7e49d46
ms.date: 06/08/2017
---


# PivotField.Subtotals Property (Excel)

Returns or sets subtotals displayed with the specified field. Valid only for nondata fields. Read/write  **Variant** .


## Syntax

 _expression_ . **Subtotals**( **_Index_** )

 _expression_ A variable that represents a **PivotField** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Optional| **Variant**|A subtotal index, as shown in the following table. If this argument is omitted, the  **Subtotals** method returns an array that contains a Boolean value for each subtotal.|

## Remarks

If an index is  **True** , the field shows that subtotal. If index 1 (Automatic) is **True** , all other values are set to **False** .



|**Index**|**Meaning**|
|:-----|:-----|
|1|Automatic|
|2|Sum|
|3|Count|
|4|Average|
|5|Max|
|6|Min|
|7|Product|
|8|Count Nums|
|9|StdDev|
|10|StdDevp|
|11|Var|
|12|Varp|
For OLAP data sources,  _Index_ can only return or be set to 1 (Automatic). The returned array always contains **True** or **False** for the first array element, and it contains **False** for all other elements. An array of element values that are all **False** indicates that there are no subtotals.


## Example

This example sets the field that contains the active cell to show Sum subtotals.


```vb
Worksheets("Sheet1").Activate 
ActiveCell.PivotField.Subtotals(2) = True
```


## See also


#### Concepts


[PivotField Object](pivotfield-object-excel.md)

