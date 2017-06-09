---
title: WorksheetFunction.Percentile_Inc Method (Excel)
keywords: vbaxl10.chm137373
f1_keywords:
- vbaxl10.chm137373
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Percentile_Inc
ms.assetid: f2c56deb-636f-7549-af70-92fc7cef3623
ms.date: 06/08/2017
---


# WorksheetFunction.Percentile_Inc Method (Excel)

Returns the k-th percentile of values in a range. You can use this function to establish a threshold of acceptance. For example, you can examine candidates who score above the 90th percentile.


## Syntax

 _expression_ . **Percentile_Inc**( **_Arg1_** , **_Arg2_** )

 _expression_ A variable that represents a **[WorksheetFunction](worksheetfunction-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Array - The array or range of data that defines relative standing.|
| _Arg2_|Required| **Double**|K - The percentile value in the range 0..1, inclusive.|

### Return Value

Double


## Remarks




- If array is empty, PERCENTILE_INC returns the #NUM! error value.
    
- If k is nonnumeric, PERCENTILE_INC returns the #VALUE! error value.
    
- If k is < 0 or if k > 1, PERCENTILE_INC returns the #NUM! error value.
    
- If k is not a multiple of 1/(n - 1), PERCENTILE_INC interpolates to determine the value at the k-th percentile.
    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

