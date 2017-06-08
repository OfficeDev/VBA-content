---
title: WorksheetFunction.Norm_S_Dist Method (Excel)
keywords: vbaxl10.chm137410
f1_keywords:
- vbaxl10.chm137410
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Norm_S_Dist
ms.assetid: ea17ac4a-82dc-ce24-0b3f-dc0452d805c6
ms.date: 06/08/2017
---


# WorksheetFunction.Norm_S_Dist Method (Excel)

Returns the standard normal cumulative distribution function. The distribution has a mean of 0 (zero) and a standard deviation of one. Use this function in place of a table of standard normal curve areas.


## Syntax

 _expression_ . **Norm_S_Dist**( **_Arg1_** )

 _expression_ A variable that represents a **[WorksheetFunction](worksheetfunction-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|Z - The value for which you want the distribution.|
| _Arg2_|Optional| **Variant**|Cumulative - A logical value that determines the form of the function. If cumulative is TRUE, NORM_S_DIST returns the cumulative distribution function; if FALSE, it returns the probability mass function.|

### Return Value

Double


## Remarks




- If z is non-numeric, NORM_S_DIST returns the #VALUE! error value.
    
- The equation for the standard normal cumulative distribution function is:
    
    
![Equation](images/abbf5ae3-a27b-4e9c-eff8-009885a4ccf2.gif)


    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

