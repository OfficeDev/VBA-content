---
title: WorksheetFunction.AverageIf Method (Excel)
keywords: vbaxl10.chm137355
f1_keywords:
- vbaxl10.chm137355
ms.prod: excel
api_name:
- Excel.WorksheetFunction.AverageIf
ms.assetid: 5409428c-ee42-8a36-42f2-f6d4ca8030d9
ms.date: 06/08/2017
---


# WorksheetFunction.AverageIf Method (Excel)

Returns the average (arithmetic mean) of all the cells in a range that meet a given criteria.


## Syntax

 _expression_ . **AverageIf**( **_Arg1_** , **_Arg2_** , **_Arg3_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Range**|One or more cells to average.|
| _Arg2_|Required| **Variant**|The criteria in the form of a number, expression, cell reference, or text that defines which cells are averaged. For example, criteria can be expressed as 32, "32", ">32", "apples", or B4.|
| _Arg3_|Optional| **Variant**|The actual set of cells to average. If omitted, range is used.|

### Return Value

Double


## Remarks




- Cells in range that contain TRUE or FALSE are ignored.
    
- If a cell in range or average_range is an empty cell, AverageIf ignores it.
    
- If a cell in criteria is empty, AverageIf treats it as a 0 value.
    
- If no cells in the range meet the criteria, AverageIf generates an error value.
    
- You can use the wildcard characters, question mark (?) and asterisk (*), in criteria. A question mark matches any single character; an asterisk matches any sequence of characters. If you want to find an actual question mark or asterisk, type a tilde (~) before the character.
    
- Average_range does not have to be the same size and shape as range. The actual cells that are averaged are determined by using the top, left cell in average_range as the beginning cell, and then including cells that correspond in size and shape to range. For example:
    

|**If range is**|**And average_range is**|**Then the actual cells evaluated are**|
|:-----|:-----|:-----|
|A1:A5|B1:B5|B1:B5|
|A1:A5|B1:B3|B1:B5|
|A1:B4|C1:D4|C1:D4|
|A1:B4|C1:C2|C1:D4|



 **Note**  The AverageIf method measures central tendency, which is the location of the center of a group of numbers in a statistical distribution. The three most common measures of central tendency are:


-  **Average** which is the arithmetic mean, and is calculated by adding a group of numbers and then dividing by the count of those numbers. For example, the average of 2, 3, 3, 5, 7, and 10 is 30 divided by 6, which is 5.
    
-  **Median** which is the middle number of a group of numbers; that is, half the numbers have values that are greater than the median, and half the numbers have values that are less than the median. For example, the median of 2, 3, 3, 5, 7, and 10 is 4.
    
-  **Mode** which is the most frequently occurring number in a group of numbers. For example, the mode of 2, 3, 3, 5, 7, and 10 is 3.
    
For a symmetrical distribution of a group of numbers, these three measures of central tendency are all the same. For a skewed distribution of a group of numbers, they can be different. 


## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

