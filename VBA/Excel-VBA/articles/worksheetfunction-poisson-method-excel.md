---
title: WorksheetFunction.Poisson Method (Excel)
keywords: vbaxl10.chm137204
f1_keywords:
- vbaxl10.chm137204
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Poisson
ms.assetid: a0c811b7-30e3-b50f-fb81-7553bb322ec1
ms.date: 06/08/2017
---


# WorksheetFunction.Poisson Method (Excel)

Returns the Poisson distribution. A common application of the Poisson distribution is predicting the number of events over a specific time, such as the number of cars arriving at a toll plaza in 1 minute.


## 


 **Important**  This function has been replaced with one or more new functions that may provide improved accuracy and whose names better reflect their usage. This function is still available for compatibility with earlier versions of Excel. However, if backward compatibility is not required, you should consider using the new functions from now on, because they more accurately describe their functionality.

For more information about the new function, see the [Poisson_Dist](worksheetfunction-poisson_dist-method-excel.md) method.


## Syntax

 _expression_ . **Poisson**( **_Arg1_** , **_Arg2_** , **_Arg3_** )

 _expression_ A variable that represents a **[WorksheetFunction](worksheetfunction-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|X - the number of events.|
| _Arg2_|Required| **Double**|Mean - the expected numeric value.|
| _Arg3_|Required| **Boolean**|Cumulative - a logical value that determines the form of the probability distribution returned. If cumulative is TRUE, POISSON returns the cumulative Poisson probability that the number of random events occurring will be between zero and x inclusive; if FALSE, it returns the Poisson probability mass function that the number of events occurring will be exactly x.|

### Return Value

Double


## Remarks




- If x is not an integer, it is truncated.
    
- If x or mean is nonnumeric, POISSON returns the #VALUE! error value.
    
- If x < 0, POISSON returns the #NUM! error value.
    
- If mean ? 0, POISSON returns the #NUM! error value.
    
- POISSON is calculated as follows. For cumulative = FALSE: 
![Formula](images/awfpois1_ZA06051232.gif)For cumulative = TRUE: 
![Formula](images/awfpois2_ZA06051233.gif)


    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

