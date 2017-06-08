---
title: WorksheetFunction.Trend Method (Excel)
keywords: vbaxl10.chm137104
f1_keywords:
- vbaxl10.chm137104
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Trend
ms.assetid: 3baae2ed-68c9-88b7-b44e-b5ea91bcbb1d
ms.date: 06/08/2017
---


# WorksheetFunction.Trend Method (Excel)

Returns values along a linear trend. Fits a straight line (using the method of least squares) to the arrays known_y's and known_x's. Returns the y-values along that line for the array of new_x's that you specify.


## Syntax

 _expression_ . **Trend**( **_Arg1_** , **_Arg2_** , **_Arg3_** , **_Arg4_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Known_y's - the set of y-values you already know in the relationship y = mx + b.|
| _Arg2_|Optional| **Variant**|Known_x's - an optional set of x-values that you may already know in the relationship y = mx + b.|
| _Arg3_|Optional| **Variant**|New_x's - new x-values for which you want TREND to return corresponding y-values.|
| _Arg4_|Optional| **Variant**|Const - a logical value specifying whether to force the constant b to equal 0.|

### Return Value

Variant


## Remarks




- If the array known_y's is in a single column, then each column of known_x's is interpreted as a separate variable.
    
- If the array known_y's is in a single row, then each row of known_x's is interpreted as a separate variable.
    

- The array known_x's can include one or more sets of variables. If only one variable is used, known_y's and known_x's can be ranges of any shape, as long as they have equal dimensions. If more than one variable is used, known_y's must be a vector (that is, a range with a height of one row or a width of one column).
    
- If known_x's is omitted, it is assumed to be the array {1,2,3,...} that is the same size as known_y's.
    

- New_x's must include a column (or row) for each independent variable, just as known_x's does. So, if known_y's is in a single column, known_x's and new_x's must have the same number of columns. If known_y's is in a single row, known_x's and new_x's must have the same number of rows.
    
- If you omit new_x's, it is assumed to be the same as known_x's.
    
- If you omit both known_x's and new_x's, they are assumed to be the array {1,2,3,...} that is the same size as known_y's.
    

- If const is TRUE or omitted, b is calculated normally.
    
- If const is FALSE, b is set equal to 0 (zero), and the m-values are adjusted so that y = mx.
    

- For information about how Microsoft Excel fits a line to data, see LINEST.
    
- You can use TREND for polynomial curve fitting by regressing against the same variable raised to different powers. For example, suppose column A contains y-values and column B contains x-values. You can enter x^2 in column C, x^3 in column D, and so on, and then regress columns B through D against column A.
    
- Formulas that return arrays must be entered as array formulas.
    
- When entering an array constant for an argument such as known_x's, use commas to separate values in the same row and semicolons to separate rows.
    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

