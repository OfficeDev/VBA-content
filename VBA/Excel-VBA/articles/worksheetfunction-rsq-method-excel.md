---
title: WorksheetFunction.RSq Method (Excel)
keywords: vbaxl10.chm137217
f1_keywords:
- vbaxl10.chm137217
ms.prod: excel
api_name:
- Excel.WorksheetFunction.RSq
ms.assetid: f6d9b270-ec48-1b53-fe96-b62dd37f1a56
ms.date: 06/08/2017
---


# WorksheetFunction.RSq Method (Excel)

Returns the square of the Pearson product moment correlation coefficient through data points in known_y's and known_x's. For more information, see PEARSON. The r-squared value can be interpreted as the proportion of the variance in y attributable to the variance in x.


## Syntax

 _expression_ . **RSq**( **_Arg1_** , **_Arg2_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Known_y's - an array or range of data points.|
| _Arg2_|Required| **Variant**|Known_x's - an array or range of data points.|

### Return Value

Double


## Remarks




-  Arguments can either be numbers or names, arrays, or references that contain numbers.
    
- Logical values and text representations of numbers that you type directly into the list of arguments are counted.
    
- If an array or reference argument contains text, logical values, or empty cells, those values are ignored; however, cells with the value zero are included.
    
- Arguments that are error values or text that cannot be translated into numbers cause errors.
    
- If known_y's and known_x's are empty or have a different number of data points, RSQ returns the #N/A error value.
    
- If known_y's and known_x's contain only 1 data point, RSQ returns the #DIV/0! error value.
    
- The equation for the Pearson product moment correlation coefficient, r, is:
![Formula](images/awfpears_ZA06051230.gif)where x and y are the sample means AVERAGE(known_x?s) and AVERAGE(known_y?s). RSQ returns r 2 , which is the square of this correlation coefficient.
    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

