---
title: WorksheetFunction.Covar Method (Excel)
keywords: vbaxl10.chm137212
f1_keywords:
- vbaxl10.chm137212
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Covar
ms.assetid: 8e08c1c6-c4c4-9088-bd2e-3ab0edc831e2
ms.date: 06/08/2017
---


# WorksheetFunction.Covar Method (Excel)

Returns covariance, the average of the products of deviations for each data point pair.


 **Important**  This function has been replaced with one or more new functions that may provide improved accuracy and whose names better reflect their usage. This function is still available for compatibility with earlier versions of Excel. However, if backward compatibility is not required, you should consider using the new functions from now on, because they more accurately describe their functionality.

For more information about the new functions, see the [Covariance_P](worksheetfunction-covar-method-excel.md) and[Covariance_S](worksheetfunction-covariance_s-method-excel.md) method.

## Syntax

 _expression_ . **Covar**( **_Arg1_** , **_Arg2_** )

 _expression_ A variable that represents a **[WorksheetFunction](worksheetfunction-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|The first cell range of integers.|
| _Arg2_|Required| **Variant**|The second cell range of integers.|

### Return Value

Double


## Remarks

 Use covariance to determine the relationship between two data sets. For example, you can examine whether greater income accompanies greater levels of education.


- The arguments must either be numbers or be names, arrays, or references that contain numbers.
    
- If an array or reference argument contains text, logical values, or empty cells, those values are ignored; however, cells with the value zero are included.
    
- If  _Arg1_ and _Arg2_ have different numbers of data points, COVAR generates an error.
    
- If either  _Arg1_ or _Arg2_ is empty, Covar generates an error.
    
- The covariance is:
![Formula](images/awfcovar_ZA06051128.gif)where x and y are the sample means AVERAGE(array1) and AVERAGE(array2), and n is the sample size. 
    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

