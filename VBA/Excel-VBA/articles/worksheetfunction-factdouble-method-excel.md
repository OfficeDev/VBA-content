---
title: WorksheetFunction.FactDouble Method (Excel)
keywords: vbaxl10.chm137292
f1_keywords:
- vbaxl10.chm137292
ms.prod: excel
api_name:
- Excel.WorksheetFunction.FactDouble
ms.assetid: 71d8d537-b06c-7614-d6d6-b6c57ed8c68f
ms.date: 06/08/2017
---


# WorksheetFunction.FactDouble Method (Excel)

Returns the double factorial of a number.


## Syntax

 _expression_ . **FactDouble**( **_Arg1_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Number - the value for which to return the double factorial. If number is not an integer, it is truncated.|

### Return Value

Double


## Remarks




- If number is nonnumeric, FACTDOUBLE returns the #VALUE! error value.
    
- If number is negative, FACTDOUBLE returns the #NUM! error value.
    
- If number is even:
![Formula](images/awffdbl1_ZA06051139.gif)


    
- If number is odd:
![Formula](images/awffdbl2_ZA06051140.gif)


    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

