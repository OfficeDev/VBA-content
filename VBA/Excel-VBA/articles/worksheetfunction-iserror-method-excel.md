---
title: WorksheetFunction.IsError Method (Excel)
keywords: vbaxl10.chm137076
f1_keywords:
- vbaxl10.chm137076
ms.prod: excel
api_name:
- Excel.WorksheetFunction.IsError
ms.assetid: 87902aa7-295b-5d0b-650e-b30b8a4084c8
ms.date: 06/08/2017
---


# WorksheetFunction.IsError Method (Excel)

Checks the type of value and returns TRUE or FALSE depending if the value refers to any error value (#N/A, #VALUE!, #REF!, #DIV/0!, #NUM!, #NAME?, or #NULL!).


## Syntax

 _expression_ . **IsError**( **_Arg1_** )

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

