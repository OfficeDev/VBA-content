---
title: WorksheetFunction.ImLog10 Method (Excel)
keywords: vbaxl10.chm137280
f1_keywords:
- vbaxl10.chm137280
ms.prod: excel
api_name:
- Excel.WorksheetFunction.ImLog10
ms.assetid: 6c391f4f-9f5c-1323-250e-2da9e055259e
ms.date: 06/08/2017
---


# WorksheetFunction.ImLog10 Method (Excel)

Returns the common logarithm (base 10) of a complex number in x + yi or x + yj text format.


## Syntax

 _expression_ . **ImLog10**( **_Arg1_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Inumber - a complex number for which you want the common logarithm.|

### Return Value

String


## Remarks




- Use COMPLEX to convert real and imaginary coefficients into a complex number.
    
- The common logarithm of a complex number can be calculated from the natural logarithm as follows:
![Formula](images/awfimlg_ZA06051160.gif)


    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

