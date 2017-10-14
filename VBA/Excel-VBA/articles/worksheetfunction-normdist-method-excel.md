---
title: WorksheetFunction.NormDist Method (Excel)
keywords: vbaxl10.chm137197
f1_keywords:
- vbaxl10.chm137197
ms.prod: excel
api_name:
- Excel.WorksheetFunction.NormDist
ms.assetid: cfc5e7e8-5723-7688-b53a-ced6bced4f58
ms.date: 06/08/2017
---


# WorksheetFunction.NormDist Method (Excel)

Returns the normal distribution for the specified mean and standard deviation. This function has a very wide range of applications in statistics, including hypothesis testing.


 **Important**  This function has been replaced with one or more new functions that may provide improved accuracy and whose names better reflect their usage. This function is still available for compatibility with earlier versions of Excel. However, if backward compatibility is not required, you should consider using the new functions from now on, because they more accurately describe their functionality.

For more information about the new function, see the [Norm_Dist](worksheetfunction-norm_dist-method-excel.md) method.

## Syntax

 _expression_ . **NormDist**( **_Arg1_** , **_Arg2_** , **_Arg3_** , **_Arg4_** )

 _expression_ A variable that represents a **[WorksheetFunction](worksheetfunction-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|X - the value for which you want the distribution.|
| _Arg2_|Required| **Double**|Mean - the arithmetic mean of the distribution.|
| _Arg3_|Required| **Double**|Standard_dev - the standard deviation of the distribution.|
| _Arg4_|Required| **Boolean**|Cumulative - a logical value that determines the form of the function. If cumulative is TRUE, NORMDIST returns the cumulative distribution function; if FALSE, it returns the probability mass function.|

### Return Value

Double


## Remarks




- If mean or standard_dev is nonnumeric, NORMDIST returns the #VALUE! error value.
    
- If standard_dev ? 0, NORMDIST returns the #NUM! error value.
    
- If mean = 0, standard_dev = 1, and cumulative = TRUE, NORMDIST returns the standard normal distribution, NORMSDIST.
    
- The equation for the normal density function (cumulative = FALSE) is:
![Formula](images/awfnrmdi_ZA06051213.gif)


    
- When cumulative = TRUE, the formula is the integral from negative infinity to x of the given formula. 
    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

