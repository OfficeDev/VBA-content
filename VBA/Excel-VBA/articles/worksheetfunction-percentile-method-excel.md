---
title: WorksheetFunction.Percentile Method (Excel)
keywords: vbaxl10.chm137232
f1_keywords:
- vbaxl10.chm137232
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Percentile
ms.assetid: a4918744-a7b1-28f9-4591-58c5ebf25c10
ms.date: 06/08/2017
---


# WorksheetFunction.Percentile Method (Excel)

Returns the k-th percentile of values in a range. You can use this function to establish a threshold of acceptance. For example, you can decide to examine candidates who score above the 90th percentile.


 **Important**  This function has been replaced with one or more new functions that may provide improved accuracy and whose names better reflect their usage. This function is still available for compatibility with earlier versions of Excel. However, if backward compatibility is not required, you should consider using the new functions from now on, because they more accurately describe their functionality.

For more information about the new functions, see the [Percentile_Inc](worksheetfunction-percentile_inc-method-excel.md) and[Percentile_Exc](worksheetfunction-percentile_exc-method-excel.md) method.

## Syntax

 _expression_ . **Percentile**( **_Arg1_** , **_Arg2_** )

 _expression_ A variable that represents a **[WorksheetFunction](worksheetfunction-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Array - the array or range of data that defines relative standing.|
| _Arg2_|Required| **Double**|K - the percentile value in the range 0..1, inclusive.|

### Return Value

Double


## Remarks




- If array is empty, PERCENTILE returns the #NUM! error value.
    
- If k is nonnumeric, PERCENTILE returns the #VALUE! error value.
    
- If k is < 0 or if k > 1, PERCENTILE returns the #NUM! error value.
    
- If k is not a multiple of 1/(n - 1), PERCENTILE interpolates to determine the value at the k-th percentile.
    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

