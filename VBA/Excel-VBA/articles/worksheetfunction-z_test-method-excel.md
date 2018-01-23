---
title: WorksheetFunction.Z_Test Method (Excel)
keywords: vbaxl10.chm137413
f1_keywords:
- vbaxl10.chm137413
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Z_Test
ms.assetid: 86c2af95-965f-f249-7775-65ff5c41785d
ms.date: 06/08/2017
---


# WorksheetFunction.Z_Test Method (Excel)

Returns the one-tailed probability-value of a z-test. For a given hypothesized population mean, Z_TEST returns the probability that the sample mean would be greater than the average of observations in the data set (array) ? that is, the observed sample mean.


## Syntax

 _expression_ . **Z_Test**( **_Arg1_** , **_Arg2_** , **_Arg3_** )

 _expression_ A variable that represents a **[WorksheetFunction](worksheetfunction-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Array is the array or range of data against which to test the hypothesized population mean.|
| _Arg2_|Required| **Double**|The value to test.|
| _Arg3_|Optional| **Variant**|Sigma - The population (known) standard deviation. If omitted, the sample standard deviation is used.|

### Return Value

Double


## Remarks




- If array is empty, Z_TEST returns the #N/A error value.
    
- Z_TEST is calculated as follows when sigma is not omitted:
![The Z_TEST calculation when sigma is not omitted](images/Z_TEST_SIGMA_ZA10391001.jpg)or when sigma is omitted: 
![The Z_TEST calculation when sigma is omitted](images/Z_TEST_ZA10391000.jpg)where x is the sample mean AVERAGE(array); s is the sample standard deviation STDEV_S(array); and n is the number of observations in the sample COUNT(array). 
    
- Z_TEST represents the probability that the sample mean would be greater than the observed value AVERAGE(array), when the underlying population mean is ? 0 . From the symmetry of the Normal distribution, if AVERAGE(array) < ?0 , Z_TEST will return a value greater than 0.5.
    
- The following Excel formula can be used to calculate the two-tailed probability that the sample mean would be further from ? 0 (in either direction) than AVERAGE(array), when the underlying population mean is ?0 : =2 * MIN(Z_TEST(array,?0 ,sigma), 1 - Z_TEST(array,?0 ,sigma)).
    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

