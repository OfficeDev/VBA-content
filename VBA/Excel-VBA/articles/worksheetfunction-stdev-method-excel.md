---
title: WorksheetFunction.StDev Method (Excel)
keywords: vbaxl10.chm137082
f1_keywords:
- vbaxl10.chm137082
ms.prod: excel
api_name:
- Excel.WorksheetFunction.StDev
ms.assetid: d401027d-672a-25a6-0d18-bcee4592e7cf
ms.date: 06/08/2017
---


# WorksheetFunction.StDev Method (Excel)

Estimates standard deviation based on a sample. The standard deviation is a measure of how widely values are dispersed from the average value (the mean).


## 


 **Important**  This function has been replaced with one or more new functions that may provide improved accuracy and whose names better reflect their usage. This function is still available for compatibility with earlier versions of Excel. However, if backward compatibility is not required, you should consider using the new functions from now on, because they more accurately describe their functionality.

For more information about the new function, see the [StDev_S](worksheetfunction-stdev_s-method-excel.md) method.


## Syntax

 _expression_ . **StDev**( **_Arg1_** , **_Arg2_** , **_Arg3_** , **_Arg4_** , **_Arg5_** , **_Arg6_** , **_Arg7_** , **_Arg8_** , **_Arg9_** , **_Arg10_** , **_Arg11_** , **_Arg12_** , **_Arg13_** , **_Arg14_** , **_Arg15_** , **_Arg16_** , **_Arg17_** , **_Arg18_** , **_Arg19_** , **_Arg20_** , **_Arg21_** , **_Arg22_** , **_Arg23_** , **_Arg24_** , **_Arg25_** , **_Arg26_** , **_Arg27_** , **_Arg28_** , **_Arg29_** , **_Arg30_** )

 _expression_ A variable that represents a **[WorksheetFunction](worksheetfunction-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1 - Arg30_|Required| **Variant**|Number1, number2, ... - 1 to 30 number arguments corresponding to a sample of a population. You can also use a single array or a reference to an array instead of arguments separated by commas.|

### Return Value

Double


## Remarks




- STDEV assumes that its arguments are a sample of the population. If your data represents the entire population, then compute the standard deviation using STDEVP.
    
- The standard deviation is calculated using the "unbiased" or "n-1" method.
    
- Arguments can either be numbers or names, arrays, or references that contain numbers.
    
- Logical values and text representations of numbers that you type directly into the list of arguments are counted.
    
- If an argument is an array or reference, only numbers in that array or reference are counted. Empty cells, logical values, text, or error values in the array or reference are ignored. 
    
- Arguments that are error values or text that cannot be translated into numbers cause errors.
    
- STDEV uses the following formula:
![Formula](images/awfstdv1_ZA06051248.gif)where x is the sample mean AVERAGE(number1,number2,?) and n is the sample size. 
    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

