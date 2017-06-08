---
title: WorksheetFunction.Covariance_S Method (Excel)
keywords: vbaxl10.chm137364
f1_keywords:
- vbaxl10.chm137364
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Covariance_S
ms.assetid: b660d4b7-80d4-3b79-f987-373f01020e6d
ms.date: 06/08/2017
---


# WorksheetFunction.Covariance_S Method (Excel)

Returns the sample covariance, the average of the products of deviations for each data point pair in two data sets.


## Syntax

 _expression_ . **Covariance_S**( **_Arg1_** , **_Arg2_** )

 _expression_ A variable that represents a **[WorksheetFunction](worksheetfunction-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Array1 - The first cell range of integers.|
| _Arg2_|Required| **Variant**|Array2 - The second cell range of integers.|

### Return Value

Double


## Remarks




- The arguments must either be numbers or be names, arrays, or references that contain numbers.
    
- If an array or reference argument contains text, logical values, or empty cells, those values are ignored; however, cells with the value zero are included.
    
- If array1 and array2 have different numbers of data points, COVARIANCE_S returns the #N/A error value.
    
- If either array1 or array2 is empty or contains only 1 data point each, COVARIANCE_S returns the #DIV/0! error value.
    



## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

