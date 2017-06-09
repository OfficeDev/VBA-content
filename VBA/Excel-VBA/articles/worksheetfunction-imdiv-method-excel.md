---
title: WorksheetFunction.ImDiv Method (Excel)
keywords: vbaxl10.chm137274
f1_keywords:
- vbaxl10.chm137274
ms.prod: excel
api_name:
- Excel.WorksheetFunction.ImDiv
ms.assetid: 6379d38c-032c-da1e-b71d-cb32f59df51d
ms.date: 06/08/2017
---


# WorksheetFunction.ImDiv Method (Excel)

Returns the quotient of two complex numbers in x + yi or x + yj text format.


## Syntax

 _expression_ . **ImDiv**( **_Arg1_** , **_Arg2_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Inumber1 - the complex numerator or dividend.|
| _Arg2_|Required| **Variant**|Inumber2 - the complex denominator or divisor.|

### Return Value

String


## Remarks




- Use COMPLEX to convert real and imaginary coefficients into a complex number.
    
- The quotient of two complex numbers is:
![Formula](images/awfimdiv_ZA06051158.gif)


    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

