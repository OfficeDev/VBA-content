---
title: WorksheetFunction.MInverse Method (Excel)
keywords: vbaxl10.chm137138
f1_keywords:
- vbaxl10.chm137138
ms.prod: excel
api_name:
- Excel.WorksheetFunction.MInverse
ms.assetid: ff41fb08-8c25-f84c-dbca-ecfe4687359e
ms.date: 06/08/2017
---


# WorksheetFunction.MInverse Method (Excel)

Returns the inverse matrix for the matrix stored in an array.


## Syntax

 _expression_ . **MInverse**( **_Arg1_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Array - a numeric array with an equal number of rows and columns.|

### Return Value

Variant


## Remarks




- The size of the array must not exceed 52 columns by 52 rows. If it does, the function returns a #VALUE! error.
    
- Array can be given as a cell range, such as A1:C3; as an array constant, such as {1,2,3;4,5,6;7,8,9}; or as a name for either of these. 
    
- If any cells in array are empty or contain text, MINVERSE returns the #VALUE! error value. 
    
- MINVERSE also returns the #VALUE! error value if array does not have an equal number of rows and columns. 
    
- Formulas that return arrays must be entered as array formulas.
    
- Inverse matrices, like determinants, are generally used for solving systems of mathematical equations involving several variables. The product of a matrix and its inverse is the identity matrix ? the square array in which the diagonal values equal 1, and all other values equal 0.
    
- As an example of how a two-row, two-column matrix is calculated, suppose that the range A1:B2 contains the letters a, b, c, and d that represent any four numbers. The following table shows the inverse of the matrix A1:B2.
    

|****|**Column A**|**Column B**|
|:-----|:-----|:-----|
|Row 1|d/(a*d-b*c)|b/(b*c-a*d)|
|Row 2|c/(b*c-a*d)|a/(a*d-b*c)|

- MINVERSE is calculated with an accuracy of approximately 16 digits, which may lead to a small numeric error when the cancellation is not complete.
    
- Some square matrices cannot be inverted and will return the #NUM! error value with MINVERSE. The determinant for a noninvertable matrix is 0.
    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

