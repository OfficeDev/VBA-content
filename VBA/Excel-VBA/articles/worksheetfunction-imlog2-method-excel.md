---
title: WorksheetFunction.ImLog2 Method (Excel)
keywords: vbaxl10.chm137279
f1_keywords:
- vbaxl10.chm137279
ms.prod: excel
api_name:
- Excel.WorksheetFunction.ImLog2
ms.assetid: 7eb55cd5-fec2-c110-981b-81c55b241900
ms.date: 06/08/2017
---


# WorksheetFunction.ImLog2 Method (Excel)

Returns the base-2 logarithm of a complex number in x + yi or x + yj text format.


## Syntax

 _expression_ . **ImLog2**( **_Arg1_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Inumber - a complex number for which you want the base-2 logarithm.|

### Return Value

String


## Remarks




- Use COMPLEX to convert real and imaginary coefficients into a complex number.
    
- The base-2 logarithm of a complex number can be calculated from the natural logarithm as follows:
![Formula](images/awfimlg2_ZA06051161.gif)


    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

