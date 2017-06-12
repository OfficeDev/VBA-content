---
title: WorksheetFunction.Growth Method (Excel)
keywords: vbaxl10.chm137106
f1_keywords:
- vbaxl10.chm137106
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Growth
ms.assetid: ecc3ffcc-9739-860a-60a6-366ef7133a33
ms.date: 06/08/2017
---


# WorksheetFunction.Growth Method (Excel)

Calculates predicted exponential growth by using existing data. GROWTH returns the y-values for a series of new x-values that you specify by using existing x-values and y-values. You can also use the GROWTH worksheet function to fit an exponential curve to existing x-values and y-values.


## Syntax

 _expression_ . **Growth**( **_Arg1_** , **_Arg2_** , **_Arg3_** , **_Arg4_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Known_y's - the set of y-values you already know in the relationship y = b*m^x.|
| _Arg2_|Optional| **Variant**|Known_x's - an optional set of x-values that you may already know in the relationship y = b*m^x.|
| _Arg3_|Optional| **Variant**|New_x's - new x-values for which you want GROWTH to return corresponding y-values.|
| _Arg4_|Optional| **Variant**|Const - a logical value specifying whether to force the constant b to equal 1.|

### Return Value

Variant


## Remarks




- If the array known_y's is in a single column, then each column of known_x's is interpreted as a separate variable.
    
- If the array known_y's is in a single row, then each row of known_x's is interpreted as a separate variable.
    
- If any of the numbers in known_y's is 0 or negative, GROWTH returns the #NUM! error value.
    

- The array known_x's can include one or more sets of variables. If only one variable is used, known_y's and known_x's can be ranges of any shape, as long as they have equal dimensions. If more than one variable is used, known_y's must be a vector (that is, a range with a height of one row or a width of one column).
    
- If known_x's is omitted, it is assumed to be the array {1,2,3,...} that is the same size as known_y's.
    

- New_x's must include a column (or row) for each independent variable, just as known_x's does. So, if known_y's is in a single column, known_x's and new_x's must have the same number of columns. If known_y's is in a single row, known_x's and new_x's must have the same number of rows.
    
- If new_x's is omitted, it is assumed to be the same as known_x's.
    
- If both known_x's and new_x's are omitted, they are assumed to be the array {1,2,3,...} that is the same size as known_y's.
    

- If const is TRUE or omitted, b is calculated normally.
    
- If const is FALSE, b is set equal to 1 and the m-values are adjusted so that y = m^x.
    

- Formulas that return arrays must be entered as array formulas after selecting the correct number of cells.
    
- When entering an array constant for an argument such as known_x's, use commas to separate values in the same row and semicolons to separate rows.
    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

