---
title: WorksheetFunction.MMult Method (Excel)
keywords: vbaxl10.chm137139
f1_keywords:
- vbaxl10.chm137139
ms.prod: excel
api_name:
- Excel.WorksheetFunction.MMult
ms.assetid: 8f410152-5682-2d71-007a-5fba5f884860
ms.date: 06/08/2017
---


# WorksheetFunction.MMult Method (Excel)

Returns the matrix product of two arrays. The result is an array with the same number of rows as array1 and the same number of columns as array2.


## Syntax

 _expression_ . **MMult**( **_Arg1_** , **_Arg2_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1 - Arg2_|Required| **Variant**|Array1, array2 - the arrays you want to multiply.|

### Return Value

Variant


## Remarks




- The number of columns in array1 must be the same as the number of rows in array2, and both arrays must contain only numbers. 
    
- Array1 and array2 can be given as cell ranges, array constants, or references.
    
- MMULT returns the #VALUE! error when:
    
      - Any cells are empty or contain text.
    
  - The number of columns in array1 is different from the number of rows in array2.
    
  - The size of the resulting array is equal to or greater than a total of 5,461 cells.
    
- The matrix product array a of two arrays b and c is:
![Formula](images/awfmmult_ZA06051209.gif)where i is the row number, and j is the column number. 
    
- Formulas that return arrays must be entered as array formulas.
    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

