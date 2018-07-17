---
title: WorksheetFunction.Weekday Method (Excel)
keywords: vbaxl10.chm137115
f1_keywords:
- vbaxl10.chm137115
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Weekday
ms.assetid: dc3140ba-98bf-8e56-5440-5eba914b30bc
ms.date: 06/08/2017
---


# WorksheetFunction.Weekday Method (Excel)

Returns the day of the week corresponding to a date. The day is given as an integer, ranging from 1 (Sunday) to 7 (Saturday), by default.


## Syntax

 _expression_ . **Weekday**( **_Arg1_** , **_Arg2_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Serial_number - a sequential number that represents the date of the day you are trying to find. Dates should be entered by using the DATE function, or as results of other formulas or functions. For example, use DATE(2008,5,23) for the 23rd day of May, 2008. Problems can occur if dates are entered as text.|
| _Arg2_|Optional| **Variant**|Return_type - a number that determines the type of return value.|

### Return Value

Double


## Remarks



|**Return_type**|**Number returned**|
|:-----|:-----|
|1 or omitted|Numbers 1 (Sunday) through 7 (Saturday). Behaves like previous versions of Microsoft Excel.|
|2|Numbers 1 (Monday) through 7 (Sunday).|
|3|Numbers 0 (Monday) through 6 (Sunday).|
|11|Numbers 1 (Monday) through 7 (Sunday).|
|12|Numbers 1 (Tuesday) through 7 (Monday)|
|13|Numbers 1 (Wednesday) through 7 (Tuesday)|
|14|Numbers 1 (Thursday) through 7 (Wednesday)|
|15|Numbers 1 (Friday) through 7 (Thursday)|
|16|Numbers 1 (Saturday) through 7 (Friday)|
|17|Numbers 1 (Sunday) through 7 (Saturday)|
Microsoft Excel stores dates as sequential serial numbers so they can be used in calculations. By default, January 1, 1900 is serial number 1, and January 1, 2008 is serial number 39448 because it is 39,448 days after January 1, 1900. Microsoft Excel for the Macintosh uses a different date system as its default. 


 **Note**  Visual Basic for Applications (VBA) calculates serial dates differently than Excel. In VBA, serial number 1 is December 31, 1899, rather than January 1, 1900. 


## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

