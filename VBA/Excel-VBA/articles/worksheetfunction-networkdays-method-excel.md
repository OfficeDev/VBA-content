---
title: WorksheetFunction.NetworkDays Method (Excel)
keywords: vbaxl10.chm137348
f1_keywords:
- vbaxl10.chm137348
ms.prod: excel
api_name:
- Excel.WorksheetFunction.NetworkDays
ms.assetid: 8b00bb8c-aa5d-74a4-76af-6e86f10ee94e
ms.date: 06/08/2017
---


# WorksheetFunction.NetworkDays Method (Excel)

Returns the number of whole working days between start_date and end_date. Working days exclude weekends and any dates identified in holidays. Use NETWORKDAYS to calculate employee benefits that accrue based on the number of days worked during a specific term.


## Syntax

 _expression_ . **NetworkDays**( **_Arg1_** , **_Arg2_** , **_Arg3_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Start_date - a date that represents the start date.|
| _Arg2_|Required| **Variant**|End_date - a date that represents the end date.|
| _Arg3_|Optional| **Variant**|Holidays - an optional range of one or more dates to exclude from the working calendar, such as state and federal holidays and floating holidays. The list can be either a range of cells that contains the dates or an array constant of the serial numbers that represent the dates.|

### Return Value

Double


## Remarks


 **Important**  Dates should be entered by using the DATE function, or as results of other formulas or functions. For example, use DATE(2008,5,23) for the 23rd day of May, 2008. Problems can occur if dates are entered as text.


- Microsoft Excel stores dates as sequential serial numbers so they can be used in calculations. By default, January 1, 1900 is serial number 1, and January 1, 2008 is serial number 39448 because it is 39,448 days after January 1, 1900. Microsoft Excel for the Macintosh uses a different date system as its default.
    
     **Note**  Visual Basic for Applications (VBA) calculates serial dates differently than Excel. In VBA, serial number 1 is December 31, 1899, rather than January 1, 1900. 
- If any argument is not a valid date, NETWORKDAYS returns the #VALUE! error value.
    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

