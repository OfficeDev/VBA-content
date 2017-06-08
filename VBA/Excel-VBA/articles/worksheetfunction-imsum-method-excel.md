---
title: WorksheetFunction.ImSum Method (Excel)
keywords: vbaxl10.chm137289
f1_keywords:
- vbaxl10.chm137289
ms.prod: excel
api_name:
- Excel.WorksheetFunction.ImSum
ms.assetid: 154d2034-8933-7b20-2cae-92580ada7250
ms.date: 06/08/2017
---


# WorksheetFunction.ImSum Method (Excel)

Returns the sum of two or more complex numbers in x + yi or x + yj text format.


## Syntax

 _expression_ . **ImSum**( **_Arg1_** , **_Arg2_** , **_Arg3_** , **_Arg4_** , **_Arg5_** , **_Arg6_** , **_Arg7_** , **_Arg8_** , **_Arg9_** , **_Arg10_** , **_Arg11_** , **_Arg12_** , **_Arg13_** , **_Arg14_** , **_Arg15_** , **_Arg16_** , **_Arg17_** , **_Arg18_** , **_Arg19_** , **_Arg20_** , **_Arg21_** , **_Arg22_** , **_Arg23_** , **_Arg24_** , **_Arg25_** , **_Arg26_** , **_Arg27_** , **_Arg28_** , **_Arg29_** , **_Arg30_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1 - Arg30_|Required| **Variant**|Inumber1,inumber2,... - 1 to 29 complex numbers to add.|

### Return Value

String


## Remarks




- Use COMPLEX to convert real and imaginary coefficients into a complex number.
    
- The sum of two complex numbers is:
![Formula](images/awfimsum_ZA06051172.gif)


    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

