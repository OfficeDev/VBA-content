---
title: WorksheetFunction.Weibull Method (Excel)
keywords: vbaxl10.chm137206
f1_keywords:
- vbaxl10.chm137206
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Weibull
ms.assetid: 2636d646-d867-a66b-ceba-b180e4ae69fa
ms.date: 06/08/2017
---


# WorksheetFunction.Weibull Method (Excel)

Returns the Weibull distribution. Use this distribution in reliability analysis, such as calculating a device's mean time to failure.


## 


 **Important**  This function has been replaced with one or more new functions that may provide improved accuracy and whose names better reflect their usage. This function is still available for compatibility with earlier versions of Excel. However, if backward compatibility is not required, you should consider using the new functions from now on, because they more accurately describe their functionality.For more information about the new function, see the [Weibull_Dist](worksheetfunction-weibull_dist-method-excel.md) method.


## Syntax

 _expression_ . **Weibull**( **_Arg1_** , **_Arg2_** , **_Arg3_** , **_Arg4_** )

 _expression_ A variable that represents a **[WorksheetFunction](worksheetfunction-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|X - the value at which to evaluate the function.|
| _Arg2_|Required| **Double**|Alpha - a parameter to the distribution.|
| _Arg3_|Required| **Double**|Beta - a parameter to the distribution.|
| _Arg4_|Required| **Boolean**|Cumulative - determines the form of the function.|

### Return Value

Double


## Remarks




- If x, alpha, or beta is nonnumeric, WEIBULL returns the #VALUE! error value.
    
- If x < 0, WEIBULL returns the #NUM! error value.
    
- If alpha ? 0 or if beta ? 0, WEIBULL returns the #NUM! error value.
    
- The equation for the Weibull cumulative distribution function is:
![Formula](images/awfweib1_ZA06051261.gif)


    
- The equation for the Weibull probability density function is:
![Formula](images/awfweib2_ZA06051262.gif)


    
- When alpha = 1, WEIBULL returns the exponential distribution with:
![Formula](images/awfweib3_ZA06051263.gif)


    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

