---
title: WorksheetFunction.Covariance_P Method (Excel)
keywords: vbaxl10.chm137363
f1_keywords:
- vbaxl10.chm137363
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Covariance_P
ms.assetid: a1cc46fe-e725-3d29-d3d3-1c6a56a67abf
ms.date: 06/08/2017
---


# WorksheetFunction.Covariance_P Method (Excel)

Returns population covariance, the average of the products of deviations for each data point pair.


## Syntax

 _expression_ . **Covariance_P**( **_Arg1_** , **_Arg2_** )

 _expression_ A variable that represents a **[WorksheetFunction](worksheetfunction-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|The first cell range of integers.|
| _Arg2_|Required| **Variant**|The second cell range of integers.|

### Return Value

Double


## Remarks

 Use Covariance_P to determine the relationship between two data sets. For example, you can examine whether greater income accompanies greater levels of education:


- The arguments must either be numbers or be names, arrays, or references that contain numbers.
    
- If an array or reference argument contains text, logical values, or empty cells, those values are ignored; however, cells with the value zero are included.
    
- If  _Arg1_ and _Arg2_ have different numbers of data points, Covariance_P generates an error.
    
- If either  _Arg1_ or _Arg2_ is empty, Covariance_P generates an error.
    
- The covariance is:
![Formula](images/awfcovar_ZA06051128.gif)where x and y are the sample means AVERAGE(array1) and AVERAGE(array2), and n is the sample size. 
    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

