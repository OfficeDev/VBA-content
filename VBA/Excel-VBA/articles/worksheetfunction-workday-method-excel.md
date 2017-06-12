---
title: WorksheetFunction.WorkDay Method (Excel)
keywords: vbaxl10.chm137347
f1_keywords:
- vbaxl10.chm137347
ms.prod: excel
api_name:
- Excel.WorksheetFunction.WorkDay
ms.assetid: 358c358f-c76e-1309-4a2f-8e50f8d7e7d9
ms.date: 06/08/2017
---


# WorksheetFunction.WorkDay Method (Excel)

Returns a number that represents a date that is the indicated number of working days before or after a date (the starting date). Working days exclude weekends and any dates identified as holidays. Use WORKDAY to exclude weekends or holidays when you calculate invoice due dates, expected delivery times, or the number of days of work performed.


## Syntax

 _expression_ . **WorkDay**( **_Arg1_** , **_Arg2_** , **_Arg3_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Start_date - a date that represents the start date.|
| _Arg2_|Required| **Variant**|Days - the number of nonweekend and nonholiday days before or after start_date. A positive value for days yields a future date; a negative value yields a past date.|
| _Arg3_|Optional| **Variant**|Holidays - an optional list of one or more dates to exclude from the working calendar, such as state and federal holidays and floating holidays. The list can be either a range of cells that contain the dates or an array constant of the serial numbers that represent the dates.|

### Return Value

Double


## Remarks


 **Important**  Dates should be entered by using the DATE function, or as results of other formulas or functions. For example, use DATE(2008,5,23) for the 23rd day of May, 2008. Problems can occur if dates are entered as text .


- Microsoft Excel stores dates as sequential serial numbers so they can be used in calculations. By default, January 1, 1900 is serial number 1, and January 1, 2008 is serial number 39448 because it is 39,448 days after January 1, 1900. Microsoft Excel for the Macintosh uses a different date system as its default.
    
     **Note**  Visual Basic for Applications (VBA) calculates serial dates differently than Excel. In VBA, serial number 1 is December 31, 1899, rather than January 1, 1900. 
- If any argument is not a valid date, WORKDAY returns the #VALUE! error value.
    
- If start_date plus days yields an invalid date, WORKDAY returns the #NUM! error value.
    
- If days is not an integer, it is truncated.
    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

