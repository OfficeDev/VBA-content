---
title: WorksheetFunction.FisherInv Method (Excel)
keywords: vbaxl10.chm137188
f1_keywords:
- vbaxl10.chm137188
ms.prod: excel
api_name:
- Excel.WorksheetFunction.FisherInv
ms.assetid: bf4656e3-b79d-7fe6-917f-16afedc736fe
ms.date: 06/08/2017
---


# WorksheetFunction.FisherInv Method (Excel)

Returns the inverse of the Fisher transformation. Use this transformation when analyzing correlations between ranges or arrays of data. If y = FISHER(x), then FISHERINV(y) = x.


## Syntax

 _expression_ . **FisherInv**( **_Arg1_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|y - the value for which you want to perform the inverse of the transformation.|

### Return Value

Double


## Remarks




- If y is nonnumeric, FISHERINV returns the #VALUE! error value.
    
- The equation for the inverse of the Fisher transformation is: 
![Formula](images/awffshri_ZA06051142.gif)


    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

