---
title: WorksheetFunction.StEyx Method (Excel)
keywords: vbaxl10.chm137218
f1_keywords:
- vbaxl10.chm137218
ms.prod: excel
api_name:
- Excel.WorksheetFunction.StEyx
ms.assetid: 6a637f86-3ef6-dc6a-fe21-51693c814159
ms.date: 06/08/2017
---


# WorksheetFunction.StEyx Method (Excel)

Returns the standard error of the predicted y-value for each x in the regression. The standard error is a measure of the amount of error in the prediction of y for an individual x.


## Syntax

 _expression_ . **StEyx**( **_Arg1_** , **_Arg2_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Known_y's - an array or range of dependent data points.|
| _Arg2_|Required| **Variant**|Known_x's - an array or range of independent data points.|

### Return Value

Double


## Remarks




- Arguments can either be numbers or names, arrays, or references that contain numbers.
    
- Logical values and text representations of numbers that you type directly into the list of arguments are counted.
    
- If an array or reference argument contains text, logical values, or empty cells, those values are ignored; however, cells with the value zero are included.
    
- Arguments that are error values or text that cannot be translated into numbers cause errors.
    
- If known_y's and known_x's have a different number of data points, STEYX returns the #N/A error value.
    
- If known_y's and known_x's are empty or have less than three data points, STEYX returns the #DIV/0! error value.
    
- The equation for the standard error of the predicted y is:
![Formula](images/awfsteyx_ZA06051250.gif)where x and y are the sample means AVERAGE(known_x?s) and AVERAGE(known_y?s), and n is the sample size. 
    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

