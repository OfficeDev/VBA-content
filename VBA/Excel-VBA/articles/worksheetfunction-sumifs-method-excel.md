---
title: WorksheetFunction.SumIfs Method (Excel)
keywords: vbaxl10.chm137353
f1_keywords:
- vbaxl10.chm137353
ms.prod: excel
api_name:
- Excel.WorksheetFunction.SumIfs
ms.assetid: 02ed74ac-0402-35fa-92d3-657de7b435ea
ms.date: 06/08/2017
---


# WorksheetFunction.SumIfs Method (Excel)

Adds the cells in a range that meet multiple criteria.


## Syntax

 _expression_ . **SumIfs**( **_Arg1_** , **_Arg2_** , **_Arg3_** , **_Arg4_** , **_Arg5_** , **_Arg6_** , **_Arg7_** , **_Arg8_** , **_Arg9_** , **_Arg10_** , **_Arg11_** , **_Arg12_** , **_Arg13_** , **_Arg14_** , **_Arg15_** , **_Arg16_** , **_Arg17_** , **_Arg18_** , **_Arg19_** , **_Arg20_** , **_Arg21_** , **_Arg22_** , **_Arg23_** , **_Arg24_** , **_Arg25_** , **_Arg26_** , **_Arg27_** , **_Arg28_** , **_Arg29_** , **_Arg30_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Range**|Sum_range - the range to sum.|
| _Arg2_|Required| **Range**|Criteria_range1, criteria_range2, ? - one or more ranges in which to evaluate the associated criteria.|
| _Arg3 - Arg30_|Required| **Variant**|Criteria1, criteria2, ? - one or more criteria in the form of a number, expression, cell reference, or text that define which cells will be added. For example, criteria can be expressed as 32, "32", ">32", "apples", or B4.|

### Return Value

Double


## Remarks




- Each cell in sum_range is summed only if all of the corresponding criteria specified are true for that cell.
    
- Cells in sum_range that contain TRUE evaluate as 1; cells in sum_range that contain FALSE evaluate as 0 (zero).
    
- You can use the wildcard characters, question mark (?) and asterisk (*), in criteria. A question mark matches any single character; an asterisk matches any sequence of characters. If you want to find an actual question mark or asterisk, type a tilde (~) before the character.
    
- Each criteria_range does not have to be the same size and shape as sum_range. The actual cells that are added are determined by using the top, left cell in that criteria_range as the beginning cell, and then including cells that correspond in size and shape to sum_range. For example:
    

|**If sum_range is**|**And a criteria_range is**|**Then the actual cells evaluated are**|
|:-----|:-----|:-----|
|A1:A5|B1:B5|B1:B5|
|A1:A5|B1:B3|B1:B5|
|A1:B4|C1:D4|C1:D4|
|A1:B4|C1:C2|C1:D4|

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

