---
title: WorksheetFunction.TrimMean Method (Excel)
keywords: vbaxl10.chm137235
f1_keywords:
- vbaxl10.chm137235
ms.prod: excel
api_name:
- Excel.WorksheetFunction.TrimMean
ms.assetid: 3ba47dcd-312b-2835-c9a4-5d5fcedee71f
ms.date: 06/08/2017
---


# WorksheetFunction.TrimMean Method (Excel)

Returns the mean of the interior of a data set. TRIMMEAN calculates the mean taken by excluding a percentage of data points from the top and bottom tails of a data set. You can use this function when you wish to exclude outlying data from your analysis.


## Syntax

 _expression_ . **TrimMean**( **_Arg1_** , **_Arg2_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Array - the array or range of values to trim and average.|
| _Arg2_|Required| **Double**|Percent - the fractional number of data points to exclude from the calculation. For example, if percent = 0.2, 4 points are trimmed from a data set of 20 points (20 x 0.2): 2 from the top and 2 from the bottom of the set.|

### Return Value

Double


## Remarks




- If percent < 0 or percent > 1, TRIMMEAN returns the #NUM! error value.
    
- TRIMMEAN rounds the number of excluded data points down to the nearest multiple of 2. If percent = 0.1, 10 percent of 30 data points equals 3 points. For symmetry, TRIMMEAN excludes a single value from the top and bottom of the data set.
    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

