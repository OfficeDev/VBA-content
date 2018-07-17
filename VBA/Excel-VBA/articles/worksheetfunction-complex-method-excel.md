---
title: WorksheetFunction.Complex Method (Excel)
keywords: vbaxl10.chm137288
f1_keywords:
- vbaxl10.chm137288
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Complex
ms.assetid: 4ea29dda-7f01-1f48-5cda-f1bc7a0a52f2
ms.date: 06/08/2017
---


# WorksheetFunction.Complex Method (Excel)

Converts real and imaginary coefficients into a complex number of the form x + yi or x + yj.


## Syntax

 _expression_ . **Complex**( **_Arg1_** , **_Arg2_** , **_Arg3_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|The real coefficient of the complex number.|
| _Arg2_|Required| **Variant**|The imaginary coefficient of the complex number.|
| _Arg3_|Optional| **Variant**|The suffix for the imaginary component of the complex number. If omitted, suffix is assumed to be "i".|

### Return Value

String


## Remarks


- If  _Arg1_ is nonnumeric, Complex generates an error.
    
- If  _Arg2_ is nonnumeric, Complex generates an error.
    
- If  _Arg3_ is neither "i" nor "j", Complex generates an error.
    

 **Note**  All complex number functions accept "i" and "j" for suffix, but neither "I" nor "J". Using uppercase generates an error. All functions that accept two or more complex numbers require that all suffixes match.


## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

