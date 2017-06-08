---
title: WorksheetFunction.Floor Method (Excel)
keywords: vbaxl10.chm137189
f1_keywords:
- vbaxl10.chm137189
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Floor
ms.assetid: c35733d5-34b9-8475-197f-4f13ae1e6c1a
ms.date: 06/08/2017
---


# WorksheetFunction.Floor Method (Excel)

Rounds number down, toward zero, to the nearest multiple of significance.


 **Important**  This function has been replaced with one or more new functions that may provide improved accuracy and whose names better reflect their usage. This function is still available for compatibility with earlier versions of Excel. However, if backward compatibility is not required, you should consider using the new functions from now on, because they more accurately describe their functionality.

For more information about the new function, see the [Floor_Precise](worksheetfunction-floor_precise-method-excel.md) method.

## Syntax

 _expression_ . **Floor**( **_Arg1_** , **_Arg2_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|Number - the numeric value you want to round.|
| _Arg2_|Required| **Double**|Significance - the multiple to which you want to round.|

### Return Value

Double


## Remarks




- As long as the number and specified significance have the same sign, then FLOOR rounds TOWARDS zero to the nearest multiple of significance.
    
- If either argument is nonnumeric, FLOOR returns the #VALUE! error value.
    
- In Excel and Excel, Excel allows positive and negative multiples of significance with negative numbers. In those cases, if the significance is positive FLOOR rounds away from zero. Otherwise,if significance is negative FLOOR rounds towards zero.
    
- For positive numbers with negative multiples of significance, Excel and Excel returns the #NUM! error value.
    
- If number is an exact multiple of significance, no rounding occurs.
    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

