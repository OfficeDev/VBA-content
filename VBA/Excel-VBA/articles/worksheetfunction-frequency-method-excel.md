---
title: WorksheetFunction.Frequency Method (Excel)
keywords: vbaxl10.chm137172
f1_keywords:
- vbaxl10.chm137172
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Frequency
ms.assetid: e13a993f-c669-45ca-90f9-41655f01cc18
ms.date: 06/08/2017
---


# WorksheetFunction.Frequency Method (Excel)

Calculates how often values occur within a range of values, and then returns a vertical array of numbers. For example, use FREQUENCY to count the number of test scores that fall within ranges of scores. Because FREQUENCY returns an array, it must be entered as an array formula.


## Syntax

 _expression_ . **Frequency**( **_Arg1_** , **_Arg2_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Data_array - an array of or reference to a set of values for which you want to count frequencies. If data_array contains no values, FREQUENCY returns an array of zeros.|
| _Arg2_|Required| **Variant**|Bins_array - an array of or reference to intervals into which you want to group the values in data_array. If bins_array contains no values, FREQUENCY returns the number of elements in data_array.|

### Return Value

Variant


## Remarks




- FREQUENCY is entered as an array formula after you select a range of adjacent cells into which you want the returned distribution to appear.
    
- The number of elements in the returned array is one more than the number of elements in bins_array. The extra element in the returned array returns the count of any values above the highest interval. For example, when counting three ranges of values (intervals) that are entered into three cells, be sure to enter FREQUENCY into four cells for the results. The extra cell returns the number of values in data_array that are greater than the third interval value.
    
- FREQUENCY ignores blank cells and text.
    
- Formulas that return arrays must be entered as array formulas.
    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

