---
title: WorksheetFunction.ZTest Method (Excel)
keywords: vbaxl10.chm137228
f1_keywords:
- vbaxl10.chm137228
ms.prod: excel
api_name:
- Excel.WorksheetFunction.ZTest
ms.assetid: 24d85668-2502-14b5-73b7-24a5dae7c332
ms.date: 06/08/2017
---


# WorksheetFunction.ZTest Method (Excel)

Returns the one-tailed probability-value of a z-test. For a given hypothesized population mean, ZTEST returns the probability that the sample mean would be greater than the average of observations in the data set (array) ? that is, the observed sample mean.


## 


 **Important**  This function has been replaced with one or more new functions that may provide improved accuracy and whose names better reflect their usage. This function is still available for compatibility with earlier versions of Excel. However, if backward compatibility is not required, you should consider using the new functions from now on, because they more accurately describe their functionality.For more information about the new function, see the [Z_Test](worksheetfunction-z_test-method-excel.md) method.


## Syntax

 _expression_ . **ZTest**( **_Arg1_** , **_Arg2_** , **_Arg3_** )

 _expression_ A variable that represents a **[WorksheetFunction](worksheetfunction-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Array is the array or range of data against which to test the hypothesized population mean.|
| _Arg2_|Required| **Double**| The value to test.|
| _Arg3_|Optional| **Variant**|Sigma - the population (known) standard deviation. If omitted, the sample standard deviation is used.|

### Return Value

Double


## Remarks




- If array is empty, ZTEST returns the #N/A error value.
    
- ZTEST is calculated as follows when sigma is not omitted:
![Formula](images/awfztest_ZA06051270.gif)or when sigma is omitted: 
![Formula](images/awfztsta_ZA06054798.gif)where x is the sample mean AVERAGE(array); s is the sample standard deviation STDEV(array); and n is the number of observations in the sample COUNT(array). 
    
- ZTEST represents the probability that the sample mean would be greater than the observed value AVERAGE(array), when the underlying population mean is ? 0 . From the symmetry of the Normal distribution, if AVERAGE(array) < ?0 , ZTEST will return a value greater than 0.5.
    
- The following Excel formula can be used to calculate the two-tailed probability that the sample mean would be further from ? 0 (in either direction) than AVERAGE(array), when the underlying population mean is ?0 : =2 * MIN(ZTEST(array,?0 ,sigma), 1 - ZTEST(array,?0 ,sigma)).
    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

