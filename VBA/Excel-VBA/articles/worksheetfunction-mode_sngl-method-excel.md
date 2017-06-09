---
title: WorksheetFunction.Mode_Sngl Method (Excel)
keywords: vbaxl10.chm137369
f1_keywords:
- vbaxl10.chm137369
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Mode_Sngl
ms.assetid: d9e3139a-8b81-69b9-11cc-93cc0357cd51
ms.date: 06/08/2017
---


# WorksheetFunction.Mode_Sngl Method (Excel)

Returns the most frequently occurring, or repetitive, value in an array or range of data.


## Syntax

 _expression_ . **Mode_Sngl**( **_Arg1_** , **_Arg2_** , **_Arg3_** , **_Arg4_** , **_Arg5_** , **_Arg6_** , **_Arg7_** , **_Arg8_** , **_Arg9_** , **_Arg10_** , **_Arg11_** , **_Arg12_** , **_Arg13_** , **_Arg14_** , **_Arg15_** , **_Arg16_** , **_Arg17_** , **_Arg18_** , **_Arg19_** , **_Arg20_** , **_Arg21_** , **_Arg22_** , **_Arg23_** , **_Arg24_** , **_Arg25_** , **_Arg26_** , **_Arg27_** , **_Arg28_** , **_Arg29_** , **_Arg30_** )

 _expression_ A variable that represents a **[WorksheetFunction](worksheetfunction-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1 - Arg30_|Required| **Variant**|Number1, number2, ... - 1 to 30 arguments for which you want to calculate the mode. You can also use a single array or a reference to an array instead of arguments separated by commas (,).|

### Return Value

Double


## Remarks


- Arguments can either be numbers or names, arrays, or references that contain numbers.
    
- If an array or reference argument contains text, logical values, or empty cells, those values are ignored; however, cells with the value zero are included.
    
- Arguments that are error values or text that cannot be translated into numbers cause errors.
    
- If the data set contains no duplicate data points, MODE_SNGL returns the #N/A error value.
    

 **Note**  The MODE_SNGL function measures central tendency, which is the location of the center of a group of numbers in a statistical distribution. The three most common measures of central tendency are:


-  **Average** The arithmetic mean, and is calculated by adding a group of numbers and then dividing by the count of those numbers. For example, the average of 2, 3, 3, 5, 7, and 10 is 30 divided by 6, which is 5.
    
-  **Median** The middle number of a group of numbers; that is, half the numbers have values that are greater than the median, and half the numbers have values that are less than the median. For example, the median of 2, 3, 3, 5, 7, and 10 is 4.
    
-  **Mode** The most frequently occurring number in a group of numbers. For example, the mode of 2, 3, 3, 5, 7, and 10 is 3.
    
For a symmetrical distribution of a group of numbers, these three measures of central tendency are all the same. For a skewed distribution of a group of numbers, they can be different. 


## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

