---
title: WorksheetFunction.IfError Method (Excel)
keywords: vbaxl10.chm137357
f1_keywords:
- vbaxl10.chm137357
ms.prod: excel
api_name:
- Excel.WorksheetFunction.IfError
ms.assetid: 864812c0-990e-2e99-3c3b-05fe5210cf16
ms.date: 06/08/2017
---


# WorksheetFunction.IfError Method (Excel)

Returns a value you specify if a formula evaluates to an error; otherwise, returns the result of the formula. Use the IFERROR function to trap and handle errors in a formula.


## Syntax

 _expression_ . **IfError**( **_Arg1_** , **_Arg2_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Value - the argument that is checked for an error.|
| _Arg2_|Required| **Variant**|Value_if_error - the value to return if the formula evaluates to an error. The following error types are evaluated: #N/A, #VALUE!, #REF!, #DIV/0!, #NUM!, #NAME?, or #NULL!.|

### Return Value

Variant


## Remarks




- If value or value_if_error is an empty cell, IFERROR treats it as an empty string value ("").
    
- If value is an array formula, IFERROR returns an array of results for each cell in the range specified in value. See the second example below.
    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

