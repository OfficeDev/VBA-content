---
title: WorksheetFunction.Count Method (Excel)
keywords: vbaxl10.chm137074
f1_keywords:
- vbaxl10.chm137074
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Count
ms.assetid: e64d9378-c1ae-4800-092b-cbdfb9c80c3a
ms.date: 06/08/2017
---


# WorksheetFunction.Count Method (Excel)

Counts the number of cells that contain numbers and counts numbers within the list of arguments.


## Syntax

 _expression_ . **Count**( **_Arg1_** , **_Arg2_** , **_Arg3_** , **_Arg4_** , **_Arg5_** , **_Arg6_** , **_Arg7_** , **_Arg8_** , **_Arg9_** , **_Arg10_** , **_Arg11_** , **_Arg12_** , **_Arg13_** , **_Arg14_** , **_Arg15_** , **_Arg16_** , **_Arg17_** , **_Arg18_** , **_Arg19_** , **_Arg20_** , **_Arg21_** , **_Arg22_** , **_Arg23_** , **_Arg24_** , **_Arg25_** , **_Arg26_** , **_Arg27_** , **_Arg28_** , **_Arg29_** , **_Arg30_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1 - Arg30_|Required| **Variant**|1 to 30 arguments that can contain or refer to a variety of different types of data, but only numbers are counted.|

### Return Value

Double


## Remarks

 Use Count to get the number of entries in a number field that is in a range or array of numbers.


- Arguments that are numbers, dates, or text representation of numbers are counted.
    
- Logical values and text representations of numbers that you type directly into the list of arguments are counted.
    
- Arguments that are error values or text that cannot be translated into numbers are ignored.
    
- If an argument is an array or reference, only numbers in that array or reference are counted. Empty cells, logical values, text, or error values in the array or reference are ignored. 
    
- If you want to count logical values, text, or error values, use the CountA function.
    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

