---
title: WorksheetFunction.Quartile_Exc Method (Excel)
keywords: vbaxl10.chm137377
f1_keywords:
- vbaxl10.chm137377
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Quartile_Exc
ms.assetid: 2b33be15-7d3c-d8be-aae1-de100de8083c
ms.date: 06/08/2017
---


# WorksheetFunction.Quartile_Exc Method (Excel)

Returns the quartile of the data set, based on percentile values from 0..1, exclusive.


## Syntax

 _expression_ . **Quartile_Exc**( **_Arg1_** , **_Arg2_** )

 _expression_ A variable that represents a **[WorksheetFunction](worksheetfunction-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Array - The array or cell range of numeric values for which you want the quartile value.|
| _Arg2_|Required| **Double**|Quart - The value to return.|

### Return Value

Double


## Remarks




- If array is empty, QUARTILE_EXC returns the #NUM! error value.
    
- If quart is not an integer, it is truncated. 
    
- If quart ? 0 or if quart ? 4, QUARTILE_EXC returns the #NUM! error value.
    
- MIN, MEDIAN, and MAX return the same value as QUARTILE_EXC when quart is equal to 0 (zero), 2, and 4, respectively.
    



## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

