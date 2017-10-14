---
title: PivotTable.GetData Method (Excel)
keywords: vbaxl10.chm235110
f1_keywords:
- vbaxl10.chm235110
ms.prod: excel
api_name:
- Excel.PivotTable.GetData
ms.assetid: c3b88918-c515-a976-5f2e-107b981ac76f
ms.date: 06/08/2017
---


# PivotTable.GetData Method (Excel)

Returns the value for the a data filed in a PivotTable.


## Syntax

 _expression_ . **GetData**( **_Name_** )

 _expression_ A variable that represents a **PivotTable** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|Describes a single cell in the PivotTable report, using syntax similar to the  **[PivotSelect](pivottable-pivotselect-method-excel.md)** method or the PivotTable report references in calculated item formulas.|

### Return Value

Double


## Example

This example shows the sum of revenues for apples in January (Data field = Revenue, Product = Apples, Month = January).


```vb
Msgbox ActiveSheet.PivotTables(1) _ 
 .GetData("'Sum of Revenue' Apples January")
```


## See also


#### Concepts


[PivotTable Object](pivottable-object-excel.md)

