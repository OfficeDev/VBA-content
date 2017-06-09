---
title: WorksheetFunction.LogNormDist Method (Excel)
keywords: vbaxl10.chm137194
f1_keywords:
- vbaxl10.chm137194
ms.prod: excel
api_name:
- Excel.WorksheetFunction.LogNormDist
ms.assetid: 93f8135e-4967-5708-1372-0c27a0d8be12
ms.date: 06/08/2017
---


# WorksheetFunction.LogNormDist Method (Excel)

Returns the cumulative lognormal distribution of x, where ln(x) is normally distributed with parameters mean and standard_dev. Use this function to analyze data that has been logarithmically transformed.


 **Important**  This function has been replaced with one or more new functions that may provide improved accuracy and whose names better reflect their usage. This function is still available for compatibility with earlier versions of Excel. However, if backward compatibility is not required, you should consider using the new functions from now on, because they more accurately describe their functionality.

For more information about the new functions, see the [LogNorm_Dist](worksheetfunction-lognorm_dist-method-excel.md) method.

## Syntax

 _expression_ . **LogNormDist**( **_Arg1_** , **_Arg2_** , **_Arg3_** )

 _expression_ A variable that represents a **[WorksheetFunction](worksheetfunction-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|X - the value at which to evaluate the function.|
| _Arg2_|Required| **Double**|Mean - the mean of ln(x).|
| _Arg3_|Required| **Double**|Standard_dev - the standard deviation of ln(x).|

### Return Value

Double


## Remarks




- If any argument is nonnumeric, LOGNORMDIST returns the #VALUE! error value.
    
- If x ? 0 or if standard_dev ? 0, LOGNORMDIST returns the #NUM! error value.
    
- The equation for the lognormal cumulative distribution function is:
![Formula](images/awflgnmd_ZA06051179.gif)


    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

