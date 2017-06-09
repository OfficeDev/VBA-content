---
title: WorksheetFunction.Quartile_Inc Method (Excel)
keywords: vbaxl10.chm137378
f1_keywords:
- vbaxl10.chm137378
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Quartile_Inc
ms.assetid: 7febaae3-28f7-5bdb-0c20-f47dfd3c4227
ms.date: 06/08/2017
---


# WorksheetFunction.Quartile_Inc Method (Excel)

Returns the quartile of a data set based on percentile values from 0..1, inclusive. Quartiles often are used in sales and survey data to divide populations into groups. For example, you can use QUARTILE_INC to find the top 25 percent of incomes in a population.


## Syntax

 _expression_ . **Quartile_Inc**( **_Arg1_** , **_Arg2_** )

 _expression_ A variable that represents a **[WorksheetFunction](worksheetfunction-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Array - the array or cell range of numeric values for which you want the quartile value.|
| _Arg2_|Required| **Double**|Quart - The value to return.|

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

- If array is empty, QUARTILE_INC returns the #NUM! error value.
    
- If quart is not an integer, it is truncated.
    
- If quart < 0 or if quart > 4, QUARTILE_INC returns the #NUM! error value.
    
- MIN, MEDIAN, and MAX return the same value as QUARTILE_INC when quart is equal to 0 (zero), 2, and 4, respectively.
    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

