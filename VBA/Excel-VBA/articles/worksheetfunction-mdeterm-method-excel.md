---
title: WorksheetFunction.MDeterm Method (Excel)
keywords: vbaxl10.chm137137
f1_keywords:
- vbaxl10.chm137137
ms.prod: excel
api_name:
- Excel.WorksheetFunction.MDeterm
ms.assetid: 90d7be4e-308a-3641-2371-819b1687df79
ms.date: 06/08/2017
---


# WorksheetFunction.MDeterm Method (Excel)

Returns the matrix determinant of an array.


## Syntax

 _expression_ . **MDeterm**( **_Arg1_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Array - a numeric array with an equal number of rows and columns.|

### Return Value

Double


## Remarks




- Array can be given as a cell range, for example, A1:C3; as an array constant, such as {1,2,3;4,5,6;7,8,9}; or as a name to either of these. 
    
- MDTERM returns the #VALUE! error when:
    
      - Any cells in array are empty or contain text.
    
  - Array does not have an equal number of rows and columns.
    
  - The size of array exceeds 73 columns by 73 rows.
    
- The matrix determinant is a number derived from the values in array. For a three-row, three-column array, A1:C3, the determinant is defined as: `MDETERM(A1:C3)` equals `A1*(B2*C3-B3*C2) + A2*(B3*C1-B1*C3) + A3*(B1*C2-B2*C1)`
    
- Matrix determinants are generally used for solving systems of mathematical equations that involve several variables.
    
- MDETERM is calculated with an accuracy of approximately 16 digits, which may lead to a small numeric error when the calculation is not complete. For example, the determinant of a singular matrix may differ from zero by 1E-16.
    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

