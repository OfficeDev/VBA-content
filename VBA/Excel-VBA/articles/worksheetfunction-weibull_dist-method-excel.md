---
title: WorksheetFunction.Weibull_Dist Method (Excel)
keywords: vbaxl10.chm137390
f1_keywords:
- vbaxl10.chm137390
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Weibull_Dist
ms.assetid: 17e5c39f-0808-2c84-a732-801fa0e342d8
ms.date: 06/08/2017
---


# WorksheetFunction.Weibull_Dist Method (Excel)

Returns the Weibull distribution. Use this distribution in reliability analysis, such as calculating the mean time to failure for a device.


## Syntax

 _expression_ . **Weibull_Dist**( **_Arg1_** , **_Arg2_** , **_Arg3_** , **_Arg4_** )

 _expression_ A variable that represents a **[WorksheetFunction](worksheetfunction-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|X - The value at which to evaluate the function.|
| _Arg2_|Required| **Double**|Alpha - A parameter to the distribution.|
| _Arg3_|Required| **Double**|Beta - A parameter to the distribution.|
| _Arg4_|Required| **Boolean**|Cumulative - Determines the form of the function.|

### Return Value

Double


## Remarks




- If x, alpha, or beta is non-numeric, WEIBULL_DIST returns the #VALUE! error value.
    
- If x < 0, WEIBULL_DIST returns the #NUM! error value.
    
- If alpha ? 0 or if beta ? 0, WEIBULL_DIST returns the #NUM! error value.
    
- The equation for the Weibull cumulative distribution function is:
![Formula](images/awfweib1_ZA06051261.gif)


    
- The equation for the Weibull probability density function is:
![Formula](images/awfweib2_ZA06051262.gif)


    
- When alpha = 1, WEIBULL_DIST returns the exponential distribution with:
![Formula](images/awfweib3_ZA06051263.gif)


    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

