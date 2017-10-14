---
title: WorksheetFunction.Percentile_Exc Method (Excel)
keywords: vbaxl10.chm137372
f1_keywords:
- vbaxl10.chm137372
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Percentile_Exc
ms.assetid: 56a7f7eb-c69c-0baa-c64b-68fb128c4861
ms.date: 06/08/2017
---


# WorksheetFunction.Percentile_Exc Method (Excel)

Returns the k-th percentile of values in a range, where k is in the range 0..1, exclusive.


## Syntax

 _expression_ . **Percentile_Exc**( **_Arg1_** , **_Arg2_** )

 _expression_ A variable that represents a **[WorksheetFunction](worksheetfunction-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Array - The array or range of data that defines relative standing.|
| _Arg2_|Required| **Double**|K - The percentile value in the range 0..1, exclusive.|

### Return Value

Double


## Remarks




- If array is empty, PERCENTILE_EXC returns the #NUM! error value
    
- If k is nonnumeric, PERCENTILE_EXC returns the #VALUE! error value. 
    
- If k is ? 0 or if k ? 1, PERCENTILE_EXC returns the #NUM! error value. 
    
- If k is not a multiple of 1/(n - 1), PERCENTILE_EXC interpolates to determine the value at the k-th percentile.
    
- PERCENTILE_EXC will interpolate when the value for the specified percentile lies between two values in the array. If it cannot interpolate for the percentile, k specified, Excel will return #NUM! error.
    



## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

