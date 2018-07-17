---
title: WorksheetFunction.PercentRank Method (Excel)
keywords: vbaxl10.chm137233
f1_keywords:
- vbaxl10.chm137233
ms.prod: excel
api_name:
- Excel.WorksheetFunction.PercentRank
ms.assetid: c8cd2c3a-0858-27fe-b764-6bc2e7e14bf8
ms.date: 06/08/2017
---


# WorksheetFunction.PercentRank Method (Excel)

Returns the rank of a value in a data set as a percentage of the data set. This function can be used to evaluate the relative standing of a value within a data set. For example, you can use PERCENTRANK to evaluate the standing of an aptitude test score among all scores for the test.


 **Important**  This function has been replaced with one or more new functions that may provide improved accuracy and whose names better reflect their usage. This function is still available for compatibility with earlier versions of Excel. However, if backward compatibility is not required, you should consider using the new functions from now on, because they more accurately describe their functionality.

For more information about the new functions, see the [Percentile_Inc](worksheetfunction-percentile_inc-method-excel.md) and[Percentile_Exc](worksheetfunction-percentile_exc-method-excel.md) method.

## Syntax

 _expression_ . **PercentRank**( **_Arg1_** , **_Arg2_** , **_Arg3_** )

 _expression_ A variable that represents a **[WorksheetFunction](worksheetfunction-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Array - the array or range of data with numeric values that defines relative standing.|
| _Arg2_|Required| **Double**|X - the value for which you want to know the rank.|
| _Arg3_|Optional| **Variant**|Significance - an optional value that identifies the number of significant digits for the returned percentage value. If omitted, PERCENTRANK uses three digits (0.xxx).|

### Return Value

Double


## Remarks




- If array is empty, PERCENTRANK returns the #NUM! error value.
    
- If significance < 1, PERCENTRANK returns the #NUM! error value.
    
- If x does not match one of the values in array, PERCENTRANK interpolates to return the correct percentage rank.
    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

