---
title: WorksheetFunction.Ceiling Method (Excel)
keywords: vbaxl10.chm137192
f1_keywords:
- vbaxl10.chm137192
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Ceiling
ms.assetid: 4994e7d0-e626-bca4-64fc-77946438f4ed
ms.date: 06/08/2017
---


# WorksheetFunction.Ceiling Method (Excel)

Returns number rounded up, away from zero, to the nearest multiple of significance.


 **Important**  This function has been replaced with one or more new functions that may provide improved accuracy and whose names better reflect their usage. This function is still available for compatibility with earlier versions of Excel. However, if backward compatibility is not required, you should consider using the new functions from now on, because they more accurately describe their functionality.

For more information about the new function, see the [Ceiling_Precise](worksheetfunction-ceiling_precise-method-excel.md) method.

## Syntax

 _expression_ . **Ceiling**( **_Arg1_** , **_Arg2_** )

 _expression_ A variable that represents a **[WorksheetFunction](worksheetfunction-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|Number - the value you want to round.|
| _Arg2_|Required| **Double**|Significance - the multiple to which you want to round.|

### Return Value

Double


## Remarks

 For example, if you want to avoid using pennies in your prices and your product is priced at $4.42, use the formula `Ceiling(4.42,0.05)` to round prices up to the nearest nickel.


- If either argument is nonnumeric,  **Ceiling** generates an error.
    
- Regardless of the sign of number, a value is rounded up when adjusted away from zero. If number is an exact multiple of significance, no rounding occurs.
    
- If number and significance have different signs,  **Ceiling** generates an error.
    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

