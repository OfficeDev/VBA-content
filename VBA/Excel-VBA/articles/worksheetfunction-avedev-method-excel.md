---
title: WorksheetFunction.AveDev Method (Excel)
keywords: vbaxl10.chm137173
f1_keywords:
- vbaxl10.chm137173
ms.prod: excel
api_name:
- Excel.WorksheetFunction.AveDev
ms.assetid: 8fb937b3-4291-e257-f96a-7e52e6714b00
ms.date: 06/08/2017
---


# WorksheetFunction.AveDev Method (Excel)

Returns the average of the absolute deviations of data points from their mean. AveDev is a measure of the variability in a data set.


## Syntax

 _expression_ . **AveDev**( **_Arg1_** , **_Arg2_** , **_Arg3_** , **_Arg4_** , **_Arg5_** , **_Arg6_** , **_Arg7_** , **_Arg8_** , **_Arg9_** , **_Arg10_** , **_Arg11_** , **_Arg12_** , **_Arg13_** , **_Arg14_** , **_Arg15_** , **_Arg16_** , **_Arg17_** , **_Arg18_** , **_Arg19_** , **_Arg20_** , **_Arg21_** , **_Arg22_** , **_Arg23_** , **_Arg24_** , **_Arg25_** , **_Arg26_** , **_Arg27_** , **_Arg28_** , **_Arg29_** , **_Arg30_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1 - Arg 30_|Required| **Variant**|1 to 30 arguments for which you want the average of the absolute deviations. You can also use a single array or a reference to an array instead of arguments separated by commas.|

### Return Value

Double


## Remarks




- AveDev is influenced by the unit of measurement in the input data.
    
- Arguments must either be numbers or be names, arrays, or references that contain numbers.
    
- Logical values and text representations of numbers that you type directly into the list of arguments are counted.
    
- If an array or reference argument contains text, logical values, or empty cells, those values are ignored; however, cells with the value zero are included.
    
- The equation for average deviation is:
![Average deviation](images/awfavedv_ZA06051110.gif)


    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

