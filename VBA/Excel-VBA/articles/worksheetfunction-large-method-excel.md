---
title: WorksheetFunction.Large Method (Excel)
keywords: vbaxl10.chm137229
f1_keywords:
- vbaxl10.chm137229
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Large
ms.assetid: d4695008-a800-955d-ce41-8988d1a869ab
ms.date: 06/08/2017
---


# WorksheetFunction.Large Method (Excel)

Returns the k-th largest value in a data set. You can use this function to select a value based on its relative standing. For example, you can use LARGE to return the highest, runner-up, or third-place score.


## Syntax

 _expression_ . **Large**( **_Arg1_** , **_Arg2_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Array - the array or range of data for which you want to determine the k-th largest value.|
| _Arg2_|Required| **Double**|K - the position (from the largest) in the array or cell range of data to return.|

### Return Value

Double


## Remarks


- If array is empty, LARGE returns the #NUM! error value.
    
- If k ? 0 or if k is greater than the number of data points, LARGE returns the #NUM! error value.
    
If n is the number of data points in a range, then LARGE(array,1) returns the largest value, and LARGE(array,n) returns the smallest value.


## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

