---
title: WorksheetFunction.FDist Method (Excel)
keywords: vbaxl10.chm137185
f1_keywords:
- vbaxl10.chm137185
ms.prod: excel
api_name:
- Excel.WorksheetFunction.FDist
ms.assetid: ddbcd66e-d85c-4f69-1ba9-138c30a3f7d4
ms.date: 06/08/2017
---


# WorksheetFunction.FDist Method (Excel)

Returns the F probability distribution. You can use this function to determine whether two data sets have different degrees of diversity. For example, you can examine the test scores of men and women entering high school and determine if the variability in the females is different from that found in the males.


 **Important**  This function has been replaced with one or more new functions that may provide improved accuracy and whose names better reflect their usage. This function is still available for compatibility with earlier versions of Excel. However, if backward compatibility is not required, you should consider using the new functions from now on, because they more accurately describe their functionality.

For more information about the new functions, see the [F_Dist_RT](worksheetfunction-f_dist_rt-method-excel.md) and[F_Dist](worksheetfunction-f_dist-method-excel.md) methods.

## Syntax

 _expression_ . **FDist**( **_Arg1_** , **_Arg2_** , **_Arg3_** )

 _expression_ A variable that represents a **[WorksheetFunction](worksheetfunction-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|X - the value at which to evaluate the function.|
| _Arg2_|Required| **Double**|Degrees_freedom1 - the numerator degrees of freedom.|
| _Arg3_|Required| **Double**|Degrees_freedom2 - the denominator degrees of freedom.|

### Return Value

Double


## Remarks




- If any argument is nonnumeric, FDIST returns the #VALUE! error value.
    
- If x is negative, FDIST returns the #NUM! error value.
    
- If degrees_freedom1 or degrees_freedom2 is not an integer, it is truncated.
    
- If degrees_freedom1 < 1 or degrees_freedom1 ? 10^10, FDIST returns the #NUM! error value.
    
- If degrees_freedom2 < 1 or degrees_freedom2 ? 10^10, FDIST returns the #NUM! error value.
    
- FDIST is calculated as FDIST=P( F>x ), where F is a random variable that has an F distribution with degrees_freedom1 and degrees_freedom2 degrees of freedom.
    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

