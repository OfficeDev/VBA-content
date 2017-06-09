---
title: WorksheetFunction.Median Method (Excel)
keywords: vbaxl10.chm137162
f1_keywords:
- vbaxl10.chm137162
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Median
ms.assetid: a2dfcdbe-2291-e346-beca-0e93c9851532
ms.date: 06/08/2017
---


# WorksheetFunction.Median Method (Excel)

Returns the median of the given numbers. The median is the number in the middle of a set of numbers.


## Syntax

 _expression_ . **Median**( **_Arg1_** , **_Arg2_** , **_Arg3_** , **_Arg4_** , **_Arg5_** , **_Arg6_** , **_Arg7_** , **_Arg8_** , **_Arg9_** , **_Arg10_** , **_Arg11_** , **_Arg12_** , **_Arg13_** , **_Arg14_** , **_Arg15_** , **_Arg16_** , **_Arg17_** , **_Arg18_** , **_Arg19_** , **_Arg20_** , **_Arg21_** , **_Arg22_** , **_Arg23_** , **_Arg24_** , **_Arg25_** , **_Arg26_** , **_Arg27_** , **_Arg28_** , **_Arg29_** , **_Arg30_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1 - Arg30_|Required| **Variant**|Number1, number2, ... - 1 to 30 numbers for which you want the median.|

### Return Value

Double


## Remarks


- If there is an even number of numbers in the set, then MEDIAN calculates the average of the two numbers in the middle. See the second formula in the example. 
    
- Arguments can either be numbers or names, arrays, or references that contain numbers.
    
- Logical values and text representations of numbers that you type directly into the list of arguments are counted.
    
- If an array or reference argument contains text, logical values, or empty cells, those values are ignored; however, cells with the value zero are included.
    
- Arguments that are error values or text that cannot be translated into numbers cause errors.
    

 **Note**  The MEDIAN function measures central tendency, which is the location of the center of a group of numbers in a statistical distribution. The three most common measures of central tendency are:


-  **Average** which is the arithmetic mean, and is calculated by adding a group of numbers and then dividing by the count of those numbers. For example, the average of 2, 3, 3, 5, 7, and 10 is 30 divided by 6, which is 5.
    
-  **Median** which is the middle number of a group of numbers; that is, half the numbers have values that are greater than the median, and half the numbers have values that are less than the median. For example, the median of 2, 3, 3, 5, 7, and 10 is 4.
    
-  **Mode** which is the most frequently occurring number in a group of numbers. For example, the mode of 2, 3, 3, 5, 7, and 10 is 3.
    
For a symmetrical distribution of a group of numbers, these three measures of central tendency are all the same. For a skewed distribution of a group of numbers, they can be different. 


## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

