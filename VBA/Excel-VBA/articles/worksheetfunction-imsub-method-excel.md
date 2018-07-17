---
title: WorksheetFunction.ImSub Method (Excel)
keywords: vbaxl10.chm137273
f1_keywords:
- vbaxl10.chm137273
ms.prod: excel
api_name:
- Excel.WorksheetFunction.ImSub
ms.assetid: bf3d6ea1-46e2-b6d3-66e0-40576db5be2f
ms.date: 06/08/2017
---


# WorksheetFunction.ImSub Method (Excel)

Returns the difference of two complex numbers in x + yi or x + yj text format.


## Syntax

 _expression_ . **ImSub**( **_Arg1_** , **_Arg2_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Inumber1 - the complex number from which to subtract inumber2.|
| _Arg2_|Required| **Variant**|Inumber2 - the complex number to subtract from inumber1.|

### Return Value

String


## Remarks




- Use COMPLEX to convert real and imaginary coefficients into a complex number.
    
- The difference of two complex numbers is:
![Formula](images/awfimsub_ZA06051171.gif)


    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

