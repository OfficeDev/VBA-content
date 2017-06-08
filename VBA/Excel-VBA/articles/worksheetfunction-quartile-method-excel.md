---
title: WorksheetFunction.Quartile Method (Excel)
keywords: vbaxl10.chm137231
f1_keywords:
- vbaxl10.chm137231
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Quartile
ms.assetid: 92893342-0ae8-a145-4b44-4236fccf2ff8
ms.date: 06/08/2017
---


# WorksheetFunction.Quartile Method (Excel)

Returns the quartile of a data set. Quartiles often are used in sales and survey data to divide populations into groups. For example, you can use QUARTILE to find the top 25 percent of incomes in a population.


## 


 **Important**  This function has been replaced with one or more new functions that may provide improved accuracy and whose names better reflect their usage. This function is still available for compatibility with earlier versions of Excel. However, if backward compatibility is not required, you should consider using the new functions from now on, because they more accurately describe their functionality.

For more information about the new functions, see the [Quartile_Inc](worksheetfunction-quartile_inc-method-excel.md) and[Quartile_Exc](worksheetfunction-quartile_exc-method-excel.md) methods.


## Syntax

 _expression_ . **Quartile**( **_Arg1_** , **_Arg2_** )

 _expression_ A variable that represents a **[WorksheetFunction](worksheetfunction-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Array - the array or cell range of numeric values for which you want the quartile value.|
| _Arg2_|Required| **Double**|Quart - indicates which value to return.|

### Return Value

Double


## Remarks



|**If quart equals**|**QUARTILE returns**|
|:-----|:-----|
|0|Minimum value|
|1|First quartile (25th percentile)|
|2|Median value (50th percentile)|
|3|Third quartile (75th percentile)|
|4|Maximum value|

- If array is empty, QUARTILE returns the #NUM! error value.
    
- If quart is not an integer, it is truncated.
    
- If quart < 0 or if quart > 4, QUARTILE returns the #NUM! error value.
    
- MIN, MEDIAN, and MAX return the same value as QUARTILE when quart is equal to 0 (zero), 2, and 4, respectively.
    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

