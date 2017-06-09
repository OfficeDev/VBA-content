---
title: WorksheetFunction.T_Test Method (Excel)
keywords: vbaxl10.chm137412
f1_keywords:
- vbaxl10.chm137412
ms.prod: excel
api_name:
- Excel.WorksheetFunction.T_Test
ms.assetid: b777b999-348c-f3a5-0a4f-6964de4122b7
ms.date: 06/08/2017
---


# WorksheetFunction.T_Test Method (Excel)

Returns the probability associated with a Student t-Test. Use T_TEST to determine whether two samples are likely to have come from the same two underlying populations that have the same mean.


## Syntax

 _expression_ . **T_Test**( **_Arg1_** , **_Arg2_** , **_Arg3_** , **_Arg4_** )

 _expression_ A variable that represents a **[WorksheetFunction](worksheetfunction-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Array1 - The first data set.|
| _Arg2_|Required| **Variant**|Array2 - The second data set.|
| _Arg3_|Required| **Double**|Tails - Specifies the number of distribution tails. If tails = 1, T_TEST uses the one-tailed distribution. If tails = 2, T_TEST uses the two-tailed distribution.|
| _Arg4_|Required| **Double**|Type - The kind of t-Test to perform.|

### Return Value

Double


## Remarks



|**If type equals**|**This test is performed**|
|:-----|:-----|
|1|Paired|
|2|Two-sample equal variance (homoscedastic)|
|3|Two-sample unequal variance (heteroscedastic)|

- If array1 and array2 have a different number of data points, and type = 1 (paired), T_TEST returns the #N/A error value.
    
- The tails and type arguments are truncated to integers.
    
- If tails or type is non-numeric, T_TEST returns the #VALUE! error value.
    
- If tails is any value other than 1 or 2, T_TEST returns the #NUM! error value.
    
- T_TEST uses the data in array1 and array2 to compute a non-negative t-statistic. If tails=1, T_TEST returns the probability of a higher value of the t-statistic under the assumption that array1 and array2 are samples from populations with the same mean. The value returned by T_TEST when tails=2 is double that returned when tails=1 and corresponds to the probability of a higher absolute value of the t-statistic under the ?same population means? assumption.
    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

