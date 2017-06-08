---
title: WorksheetFunction.BesselI Method (Excel)
keywords: vbaxl10.chm137305
f1_keywords:
- vbaxl10.chm137305
ms.prod: excel
api_name:
- Excel.WorksheetFunction.BesselI
ms.assetid: 06bce6ff-a7cb-d8c7-2d80-d9fd54f9324b
ms.date: 06/08/2017
---


# WorksheetFunction.BesselI Method (Excel)

Returns the modified Bessel function, which is equivalent to the Bessel function evaluated for purely imaginary arguments.


## Syntax

 _expression_ . **BesselI**( **_Arg1_** , **_Arg2_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|The value at which to evaluate the function.|
| _Arg2_|Required| **Variant**| The order of the Bessel function. If n is not an integer, it is truncated.|

### Return Value

Double


## Remarks




- If x is nonnumeric, BesselI returns the #VALUE! error value.
    
- If n is nonnumeric, BesselI generates an error value.
    
- If n < 0, BesselI generates an error value.
    
- The n-th order modified Bessel function of the variable x is:
![Bessel function](images/awfbesli_ZA06051111.gif)


    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

