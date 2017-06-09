---
title: WorksheetFunction.Skew Method (Excel)
keywords: vbaxl10.chm137227
f1_keywords:
- vbaxl10.chm137227
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Skew
ms.assetid: cf10325a-0cb3-4779-d792-af365a830af9
ms.date: 06/08/2017
---


# WorksheetFunction.Skew Method (Excel)

Returns the skewness of a distribution. Skewness characterizes the degree of asymmetry of a distribution around its mean. Positive skewness indicates a distribution with an asymmetric tail extending toward more positive values. Negative skewness indicates a distribution with an asymmetric tail extending toward more negative values.


## Syntax

 _expression_ . **Skew**( **_Arg1_** , **_Arg2_** , **_Arg3_** , **_Arg4_** , **_Arg5_** , **_Arg6_** , **_Arg7_** , **_Arg8_** , **_Arg9_** , **_Arg10_** , **_Arg11_** , **_Arg12_** , **_Arg13_** , **_Arg14_** , **_Arg15_** , **_Arg16_** , **_Arg17_** , **_Arg18_** , **_Arg19_** , **_Arg20_** , **_Arg21_** , **_Arg22_** , **_Arg23_** , **_Arg24_** , **_Arg25_** , **_Arg26_** , **_Arg27_** , **_Arg28_** , **_Arg29_** , **_Arg30_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1 - Arg30_|Required| **Variant**|Number1, number2 ... - 1 to 30 arguments for which you want to calculate skewness. You can also use a single array or a reference to an array instead of arguments separated by commas.|

### Return Value

Double


## Remarks




- Arguments can either be numbers or names, arrays, or references that contain numbers.
    
- Logical values and text representations of numbers that you type directly into the list of arguments are counted.
    
- If an array or reference argument contains text, logical values, or empty cells, those values are ignored; however, cells with the value zero are included.
    
- Arguments that are error values or text that cannot be translated into numbers cause errors.
    
- If there are fewer than three data points, or the sample standard deviation is zero, SKEW returns the #DIV/0! error value.
    
- The equation for skewness is defined as:
![Formula](images/awfskew_ZA06051241.gif)


    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

