---
title: WorksheetFunction.StDev_S Method (Excel)
keywords: vbaxl10.chm137381
f1_keywords:
- vbaxl10.chm137381
ms.prod: excel
api_name:
- Excel.WorksheetFunction.StDev_S
ms.assetid: 8c62edde-7978-8b75-8554-2a1a77a5f0e2
ms.date: 06/08/2017
---


# WorksheetFunction.StDev_S Method (Excel)

Estimates standard deviation based on a sample. The standard deviation is a measure of how widely values are dispersed from the average value (the mean).


## Syntax

 _expression_ . **StDev_S**( **_Arg1_** , **_Arg2_** , **_Arg3_** , **_Arg4_** , **_Arg5_** , **_Arg6_** , **_Arg7_** , **_Arg8_** , **_Arg9_** , **_Arg10_** , **_Arg11_** , **_Arg12_** , **_Arg13_** , **_Arg14_** , **_Arg15_** , **_Arg16_** , **_Arg17_** , **_Arg18_** , **_Arg19_** , **_Arg20_** , **_Arg21_** , **_Arg22_** , **_Arg23_** , **_Arg24_** , **_Arg25_** , **_Arg26_** , **_Arg27_** , **_Arg28_** , **_Arg29_** , **_Arg30_** )

 _expression_ A variable that represents a **[WorksheetFunction](worksheetfunction-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1 - Arg30_|Required| **Variant**|Number1, number2, ... - 1 to 30 number arguments corresponding to a sample of a population. You can also use a single array or a reference to an array instead of arguments separated by commas.|

### Return Value

Double


## Remarks




- STDEV_S assumes that its arguments are a sample of the population. If your data represents the entire population, then compute the standard deviation using STDEV_P.
    
- The standard deviation is calculated using the "unbiased" or "n-1" method.
    
- Arguments can either be numbers or names, arrays, or references that contain numbers.
    
- Logical values and text representations of numbers that you type directly into the list of arguments are counted.
    
- If an argument is an array or reference, only numbers in that array or reference are counted. Empty cells, logical values, text, or error values in the array or reference are ignored. 
    
- Arguments that are error values or text that cannot be translated into numbers cause errors.
    
- STDEV_S uses the following formula:
![Formula](images/awfstdv1_ZA06051248.gif)where x is the sample mean AVERAGE(number1,number2,?) and n is the sample size. 
    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

