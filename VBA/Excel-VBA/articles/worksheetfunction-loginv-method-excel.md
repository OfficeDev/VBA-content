---
title: WorksheetFunction.LogInv Method (Excel)
keywords: vbaxl10.chm137195
f1_keywords:
- vbaxl10.chm137195
ms.prod: excel
api_name:
- Excel.WorksheetFunction.LogInv
ms.assetid: 414a4e30-1225-279b-2981-bbb798338b18
ms.date: 06/08/2017
---


# WorksheetFunction.LogInv Method (Excel)

Use the lognormal distribution to analyze logarithmically transformed data.


 **Important**  This function has been replaced with one or more new functions that may provide improved accuracy and whose names better reflect their usage. This function is still available for compatibility with earlier versions of Excel. However, if backward compatibility is not required, you should consider using the new functions from now on, because they more accurately describe their functionality.For more information about the new function, see the [LogNorm_Inv](worksheetfunction-lognorm_inv-method-excel.md) method.


## Syntax

 _expression_ . **LogInv**( **_Arg1_** , **_Arg2_** , **_Arg3_** )

 _expression_ A variable that represents a **[WorksheetFunction](worksheetfunction-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|Probability - a probability associated with the lognormal distribution.|
| _Arg2_|Required| **Double**|Mean - the mean of ln(x).|
| _Arg3_|Required| **Double**|Standard_dev - the standard deviation of ln(x).|

### Return Value

Double


## Remarks




- If any argument is nonnumeric, LOGINV returns the #VALUE! error value.
    
- If probability <= 0 or probability >= 1, LOGINV returns the #NUM! error value.
    
- If standard_dev <= 0, LOGINV returns the #NUM! error value.
    
-  The inverse of the lognormal distribution function is:
![Formula](images/awflginv_ZA06051178.gif)


    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

