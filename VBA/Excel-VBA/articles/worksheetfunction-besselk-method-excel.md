---
title: WorksheetFunction.BesselK Method (Excel)
keywords: vbaxl10.chm137303
f1_keywords:
- vbaxl10.chm137303
ms.prod: excel
api_name:
- Excel.WorksheetFunction.BesselK
ms.assetid: 9b2eb52e-2b8a-3608-6410-52abccc886b3
ms.date: 06/08/2017
---


# WorksheetFunction.BesselK Method (Excel)

Returns the modified Bessel function, which is equivalent to the Bessel functions evaluated for purely imaginary arguments.


## Syntax

 _expression_ . **BesselK**( **_Arg1_** , **_Arg2_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**| the value at which to evaluate the function.|
| _Arg2_|Required| **Variant**|the order of the function. If n is not an integer, it is truncated.|

### Return Value

Double


## Remarks




- If x is nonnumeric, BesselK generates an error value.
    
- If n is nonnumeric, BesselK generates an error value.
    
- If n < 0, BesselK generates an error value.
    
- The n-th order modified Bessel function of the variable x is:
![Bessel function](images/awfbeslk_ZA06051112.gif)where Jn and Yn are the J and Y Bessel functions, respectively. 
    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

