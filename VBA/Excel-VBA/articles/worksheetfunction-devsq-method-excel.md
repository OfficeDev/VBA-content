---
title: WorksheetFunction.DevSq Method (Excel)
keywords: vbaxl10.chm137222
f1_keywords:
- vbaxl10.chm137222
ms.prod: excel
api_name:
- Excel.WorksheetFunction.DevSq
ms.assetid: 9f74f91c-f9c0-4ffb-1145-32f010bcc257
ms.date: 06/08/2017
---


# WorksheetFunction.DevSq Method (Excel)

Returns the sum of squares of deviations of data points from their sample mean.


## Syntax

 _expression_ . **DevSq**( **_Arg1_** , **_Arg2_** , **_Arg3_** , **_Arg4_** , **_Arg5_** , **_Arg6_** , **_Arg7_** , **_Arg8_** , **_Arg9_** , **_Arg10_** , **_Arg11_** , **_Arg12_** , **_Arg13_** , **_Arg14_** , **_Arg15_** , **_Arg16_** , **_Arg17_** , **_Arg18_** , **_Arg19_** , **_Arg20_** , **_Arg21_** , **_Arg22_** , **_Arg23_** , **_Arg24_** , **_Arg25_** , **_Arg26_** , **_Arg27_** , **_Arg28_** , **_Arg29_** , **_Arg30_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1 - Arg30_|Required| **Variant**|Number1, number2, ... - are 1 to 30 arguments for which you want to calculate the sum of squared deviations. You can also use a single array or a reference to an array instead of arguments separated by commas.|

### Return Value

Double


## Remarks




- Arguments can either be numbers or names, arrays, or references that contain numbers.
    
- Logical values and text representations of numbers that you type directly into the list of arguments are counted.
    
- If an array or reference argument contains text, logical values, or empty cells, those values are ignored; however, cells with the value zero are included.
    
- Arguments that are error values or text that cannot be translated into numbers cause errors.
    
- The equation for the sum of squared deviations is:
![Formula](images/awfdevsq_ZA06051133.gif)


    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

