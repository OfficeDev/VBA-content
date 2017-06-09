---
title: WorksheetFunction.Confidence_T Method (Excel)
keywords: vbaxl10.chm137360
f1_keywords:
- vbaxl10.chm137360
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Confidence_T
ms.assetid: b4e497b6-bf5a-5630-3092-d806012e0c97
ms.date: 06/08/2017
---


# WorksheetFunction.Confidence_T Method (Excel)

Returns the confidence interval for a population mean, using a Student's t distribution.


## Syntax

 _expression_ . **Confidence_T**( **_Arg1_** , **_Arg2_** , **_Arg3_** )

 _expression_ A variable that represents a **[WorksheetFunction](worksheetfunction-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|Alpha - The significance level used to compute the confidence level. The confidence level equals 100*(1 - alpha)%, or in other words, an alpha of 0.05 indicates a 95 percent confidence level.|
| _Arg2_|Required| **Double**|Standard_dev - The population standard deviation for the data range and is assumed to be known.|
| _Arg3_|Required| **Double**|Size - The sample size.|

### Return Value

Double


## Remarks




- If any argument is nonnumeric, CONFIDENCE_T returns the #VALUE! error value. 
    
- If alpha ? 0 or alpha ? 1, CONFIDENCE_T returns the #NUM! error value. 
    
- If standard_dev ? 0, CONFIDENCE_T returns the #NUM! error value. 
    
- If size is not an integer, it is truncated. 
    
- If size equals 1, CONFIDENCE_T returns #DIV/0! error value.
    
- If size equals 1, CONFIDENCE_T returns #DIV/0! error value.
    



## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

