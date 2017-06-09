---
title: WorksheetFunction.ImSqrt Method (Excel)
keywords: vbaxl10.chm137277
f1_keywords:
- vbaxl10.chm137277
ms.prod: excel
api_name:
- Excel.WorksheetFunction.ImSqrt
ms.assetid: 095ecba9-c987-8b58-f07e-d0f79436d650
ms.date: 06/08/2017
---


# WorksheetFunction.ImSqrt Method (Excel)

Returns the square root of a complex number in x + yi or x + yj text format.


## Syntax

 _expression_ . **ImSqrt**( **_Arg1_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Inumber - a complex number for which you want the square root.|

### Return Value

String


## Remarks




- Use COMPLEX to convert real and imaginary coefficients into a complex number.
    
- The square root of a complex number is:
![Formula](images/awfimsq1_ZA06051168.gif)where: 
![Formula](images/awfimsq2_ZA06051169.gif)and: 
![Formula](images/awfimsq3_ZA06051170.gif)and: 
![Formula](images/awfimar3_ZA06051155.gif)


    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

