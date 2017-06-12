---
title: WorksheetFunction.IsErr Method (Excel)
keywords: vbaxl10.chm137130
f1_keywords:
- vbaxl10.chm137130
ms.prod: excel
api_name:
- Excel.WorksheetFunction.IsErr
ms.assetid: 478cc69a-7b1f-7c08-078d-8e56c0516ccb
ms.date: 06/08/2017
---


# WorksheetFunction.IsErr Method (Excel)

Checks the type of value and returns TRUE or FALSE depending if the value refers to any error value except #N/A.


## Syntax

 _expression_ . **IsErr**( **_Arg1_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Value - the value you want tested. Value can be a blank (empty cell), error, logical, text, number, or reference value, or a name referring to any of these, that you want to test.|

### Return Value

Boolean


## Remarks




- The value arguments of the IS functions are not converted. For example, in most other functions where a number is required, the text value "19" is converted to the number 19. However, in the formula ISNUMBER("19"), "19" is not converted from a text value, and the ISNUMBER function returns FALSE.
    
- The IS functions are useful in formulas for testing the outcome of a calculation. When combined with the IF function, they provide a method for locating errors in formulas (see the following examples).
    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

