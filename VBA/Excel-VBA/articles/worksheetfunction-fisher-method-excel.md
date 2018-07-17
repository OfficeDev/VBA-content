---
title: WorksheetFunction.Fisher Method (Excel)
keywords: vbaxl10.chm137187
f1_keywords:
- vbaxl10.chm137187
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Fisher
ms.assetid: c7326a23-f9ea-76a8-d1c4-700962362cd0
ms.date: 06/08/2017
---


# WorksheetFunction.Fisher Method (Excel)

Returns the Fisher transformation at x. This transformation produces a function that is normally distributed rather than skewed. Use this function to perform hypothesis testing on the correlation coefficient.


## Syntax

 _expression_ . **Fisher**( **_Arg1_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|x - a numeric value for which you want the transformation.|

### Return Value

Double


## Remarks




- If x is nonnumeric, FISHER returns the #VALUE! error value.
    
- If x ? -1 or if x ? 1, FISHER returns the #NUM! error value.
    
- The equation for the Fisher transformation is: 
![Formula](images/awffishr_ZA06051141.gif)


    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

