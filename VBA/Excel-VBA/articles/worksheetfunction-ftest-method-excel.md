---
title: WorksheetFunction.FTest Method (Excel)
keywords: vbaxl10.chm137214
f1_keywords:
- vbaxl10.chm137214
ms.prod: excel
api_name:
- Excel.WorksheetFunction.FTest
ms.assetid: e1f01a38-2957-a97c-d84b-f6efdec88631
ms.date: 06/08/2017
---


# WorksheetFunction.FTest Method (Excel)

Returns the result of an F-test. An F-test returns the two-tailed probability that the variances in array1 and array2 are not significantly different. Use this function to determine whether two samples have different variances. For example, given test scores from public and private schools, you can test whether these schools have different levels of test score diversity.


 **Important**  This function has been replaced with one or more new functions that may provide improved accuracy and whose names better reflect their usage. This function is still available for compatibility with earlier versions of Excel. However, if backward compatibility is not required, you should consider using the new functions from now on, because they more accurately describe their functionality.

For more information about the new function, see the [F_Test](worksheetfunction-f_test-method-excel.md) method.

## Syntax

 _expression_ . **FTest**( **_Arg1_** , **_Arg2_** )

 _expression_ A variable that represents a **[WorksheetFunction](worksheetfunction-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Array1 - the first array or range of data.|
| _Arg2_|Required| **Variant**|Array2 - the second array or range of data.|

### Return Value

Double


## Remarks




- The arguments must be either numbers or names, arrays, or references that contain numbers.
    
- If an array or reference argument contains text, logical values, or empty cells, those values are ignored; however, cells with the value zero are included.
    
- If the number of data points in array1 or array2 is less than 2, or if the variance of array1 or array2 is zero, FTEST returns the #DIV/0! error value.
    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

