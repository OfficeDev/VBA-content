---
title: WorksheetFunction.AverageIfs Method (Excel)
keywords: vbaxl10.chm137356
f1_keywords:
- vbaxl10.chm137356
ms.prod: excel
api_name:
- Excel.WorksheetFunction.AverageIfs
ms.assetid: ec1071d7-c36d-4894-dee9-6b5423f13c0b
ms.date: 06/08/2017
---


# WorksheetFunction.AverageIfs Method (Excel)

Returns the average (arithmetic mean) of all cells that meet multiple criteria.


## Syntax

 _expression_ . **AverageIfs**( **_Arg1_** , **_Arg2_** , **_Arg3_** , **_Arg4_** , **_Arg5_** , **_Arg6_** , **_Arg7_** , **_Arg8_** , **_Arg9_** , **_Arg10_** , **_Arg11_** , **_Arg12_** , **_Arg13_** , **_Arg14_** , **_Arg15_** , **_Arg16_** , **_Arg17_** , **_Arg18_** , **_Arg19_** , **_Arg20_** , **_Arg21_** , **_Arg22_** , **_Arg23_** , **_Arg24_** , **_Arg25_** , **_Arg26_** , **_Arg27_** , **_Arg28_** , **_Arg29_** , **_Arg30_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1 - Arg30_|Required| **Range**|One or more ranges in which to evaluate the associated criteria.|

### Return Value

Double


## Remarks


- If a cell in average_range is an empty cell, AverageIfs ignores it.
    
- If a cell in a criteria range is empty, AverageIfs treats it as a 0 value.
    
- Cells in range that contain TRUE evaluate as 1; cells in range that contain FALSE evaluate as 0 (zero).
    
- Each cell in average_range is used in the average calculation only if all of the corresponding criteria specified are true for that cell.
    
- If cells in average_range are empty or contain text values that cannot be translated into numbers, AverageIfs generates an error.
    
- If there are no cells that meet all the criteria, AverageIfs generates an error value.
    
- You can use the wildcard characters, question mark (?) and asterisk (*), in criteria. A question mark matches any single character; an asterisk matches any sequence of characters. If you want to find an actual question mark or asterisk, type a tilde (~) before the character.
    
- Each criteria_range does not have to be the same size and shape as average_range. The actual cells that are averaged are determined by using the top, left cell in that criteria_range as the beginning cell, and then including cells that correspond in size and shape to range. For example:
    

|**If average_range is**|**And the criteria_range is**|**Then the actual cells evaluated are**|
|:-----|:-----|:-----|
|A1:A5|B1:B5|B1:B5|
|A1:A5|B1:B3|B1:B5|
|A1:B4|C1:D4|C1:D4|
|A1:B4|C1:C2|C1:D4|

 **Note**  The AverageIfs function measures central tendency, which is the location of the center of a group of numbers in a statistical distribution. The three most common measures of central tendency are:


-  **Average** which is the arithmetic mean, and is calculated by adding a group of numbers and then dividing by the count of those numbers. For example, the average of 2, 3, 3, 5, 7, and 10 is 30 divided by 6, which is 5.
    
-  **Median** which is the middle number of a group of numbers; that is, half the numbers have values that are greater than the median, and half the numbers have values that are less than the median. For example, the median of 2, 3, 3, 5, 7, and 10 is 4.
    
-  **Mode** which is the most frequently occurring number in a group of numbers. For example, the mode of 2, 3, 3, 5, 7, and 10 is 3.
    
For a symmetrical distribution of a group of numbers, these three measures of central tendency are all the same. For a skewed distribution of a group of numbers, they can be different. 


## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

