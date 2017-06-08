---
title: WorksheetFunction.Mode Method (Excel)
keywords: vbaxl10.chm137234
f1_keywords:
- vbaxl10.chm137234
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Mode
ms.assetid: 1e26c837-159a-63cb-17ab-43bfb788a539
ms.date: 06/08/2017
---


# WorksheetFunction.Mode Method (Excel)

Returns the most frequently occurring, or repetitive, value in an array or range of data. 


 **Important**  This function has been replaced with one or more new functions that may provide improved accuracy and whose names better reflect their usage. This function is still available for compatibility with earlier versions of Excel. However, if backward compatibility is not required, you should consider using the new functions from now on, because they more accurately describe their functionality.

For more information about the new functions, see the [Mode_Sngl](worksheetfunction-mode_sngl-method-excel.md) and[Mode_Mult](worksheetfunction-mode_mult-method-excel.md) methods.

## Syntax

 _expression_ . **Mode**( **_Arg1_** , **_Arg2_** , **_Arg3_** , **_Arg4_** , **_Arg5_** , **_Arg6_** , **_Arg7_** , **_Arg8_** , **_Arg9_** , **_Arg10_** , **_Arg11_** , **_Arg12_** , **_Arg13_** , **_Arg14_** , **_Arg15_** , **_Arg16_** , **_Arg17_** , **_Arg18_** , **_Arg19_** , **_Arg20_** , **_Arg21_** , **_Arg22_** , **_Arg23_** , **_Arg24_** , **_Arg25_** , **_Arg26_** , **_Arg27_** , **_Arg28_** , **_Arg29_** , **_Arg30_** )

 _expression_ A variable that represents a **[WorksheetFunction](worksheetfunction-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1 - Arg30_|Required| **Variant**|Number1, number2, ... - 1 to 30 arguments for which you want to calculate the mode. You can also use a single array or a reference to an array instead of arguments separated by commas.|

### Return Value

Double


## Remarks


- Arguments can either be numbers or names, arrays, or references that contain numbers.
    
- If an array or reference argument contains text, logical values, or empty cells, those values are ignored; however, cells with the value zero are included.
    
- Arguments that are error values or text that cannot be translated into numbers cause errors.
    
- If the data set contains no duplicate data points, MODE returns the #N/A error value.
    

 **Note**  The MODE function measures central tendency, which is the location of the center of a group of numbers in a statistical distribution. The three most common measures of central tendency are:


-  **Average** which is the arithmetic mean, and is calculated by adding a group of numbers and then dividing by the count of those numbers. For example, the average of 2, 3, 3, 5, 7, and 10 is 30 divided by 6, which is 5.
    
-  **Median** which is the middle number of a group of numbers; that is, half the numbers have values that are greater than the median, and half the numbers have values that are less than the median. For example, the median of 2, 3, 3, 5, 7, and 10 is 4.
    
-  **Mode** which is the most frequently occurring number in a group of numbers. For example, the mode of 2, 3, 3, 5, 7, and 10 is 3.
    
For a symmetrical distribution of a group of numbers, these three measures of central tendency are all the same. For a skewed distribution of a group of numbers, they can be different. 


## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

