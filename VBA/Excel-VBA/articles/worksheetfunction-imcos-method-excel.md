---
title: WorksheetFunction.ImCos Method (Excel)
keywords: vbaxl10.chm137282
f1_keywords:
- vbaxl10.chm137282
ms.prod: excel
api_name:
- Excel.WorksheetFunction.ImCos
ms.assetid: 959ac671-64e4-ac72-9421-d7074bd5d4a8
ms.date: 06/08/2017
---


# WorksheetFunction.ImCos Method (Excel)

Returns the cosine of a complex number in x + yi or x + yj text format.


## Syntax

 _expression_ . **ImCos**( **_Arg1_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Inumber - a complex number for which you want the cosine.|

### Return Value

String


## Remarks




- Use COMPLEX to convert real and imaginary coefficients into a complex number.
    
- If inumber is a logical value, IMCOS returns the #VALUE! error value.
    
- The cosine of a complex number is:
![Formula](images/awfimcos_ZA06051157.gif)


    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

