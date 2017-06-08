---
title: WorksheetFunction.TTest Method (Excel)
keywords: vbaxl10.chm137220
f1_keywords:
- vbaxl10.chm137220
ms.prod: excel
api_name:
- Excel.WorksheetFunction.TTest
ms.assetid: 3153c88c-aa22-230f-e602-03b902830c54
ms.date: 06/08/2017
---


# WorksheetFunction.TTest Method (Excel)

Returns the probability associated with a Student's t-Test. Use TTEST to determine whether two samples are likely to have come from the same two underlying populations that have the same mean.


## 


 **Important**  This function has been replaced with one or more new functions that may provide improved accuracy and whose names better reflect their usage. This function is still available for compatibility with earlier versions of Excel. However, if backward compatibility is not required, you should consider using the new functions from now on, because they more accurately describe their functionality.For more information about the new function, see the [T_Test](worksheetfunction-t_test-method-excel.md) method.


## Syntax

 _expression_ . **TTest**( **_Arg1_** , **_Arg2_** , **_Arg3_** , **_Arg4_** )

 _expression_ A variable that represents a **[WorksheetFunction](worksheetfunction-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Array1 - the first data set.|
| _Arg2_|Required| **Variant**|Array2 - the second data set.|
| _Arg3_|Required| **Double**|Tails - specifies the number of distribution tails. If tails = 1, TTEST uses the one-tailed distribution. If tails = 2, TTEST uses the two-tailed distribution.|
| _Arg4_|Required| **Double**|Type - the kind of t-Test to perform.|

### Return Value

Double


## Remarks



|**If type equals**|**This test is performed**|
|:-----|:-----|
|1|Paired|
|2|Two-sample equal variance (homoscedastic)|
|3|Two-sample unequal variance (heteroscedastic)|

- If array1 and array2 have a different number of data points, and type = 1 (paired), TTEST returns the #N/A error value.
    
- The tails and type arguments are truncated to integers.
    
- If tails or type is nonnumeric, TTEST returns the #VALUE! error value.
    
- If tails is any value other than 1 or 2, TTEST returns the #NUM! error value.
    
- TTEST uses the data in array1 and array2 to compute a non-negative t-statistic. If tails=1, TTEST returns the probability of a higher value of the t-statistic under the assumption that array1 and array2 are samples from populations with the same mean. The value returned by TTEST when tails=2 is double that returned when tails=1 and corresponds to the probability of a higher absolute value of the t-statistic under the ?same population means? assumption.
    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

