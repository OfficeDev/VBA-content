---
title: WorksheetFunction.YearFrac Method (Excel)
keywords: vbaxl10.chm137327
f1_keywords:
- vbaxl10.chm137327
ms.prod: excel
api_name:
- Excel.WorksheetFunction.YearFrac
ms.assetid: 01c2b4c9-5a9b-6fa1-c189-7210a31583d1
ms.date: 06/08/2017
---


# WorksheetFunction.YearFrac Method (Excel)

Calculates the fraction of the year represented by the number of whole days between two dates (the start_date and the end_date). Use the YEARFRAC worksheet function to identify the proportion of a whole year's benefits or obligations to assign to a specific term.


## Syntax

 _expression_ . **YearFrac**( **_Arg1_** , **_Arg2_** , **_Arg3_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Start_date - a date that represents the start date.|
| _Arg2_|Required| **Variant**|End_date - a date that represents the end date.|
| _Arg3_|Optional| **Variant**|Basis - the type of day count basis to use.|

### Return Value

Double


## Remarks


 **Important**  Dates should be entered by using the DATE function, or as results of other formulas or functions. For example, use DATE(2008,5,23) for the 23rd day of May, 2008. Problems can occur if dates are entered as text.



|**Basis**|**Day count basis**|
|:-----|:-----|
|0 or omitted|US (NASD) 30/360|
|1|Actual/actual|
|2|Actual/360|
|3|Actual/365|
|4|European 30/360|

- Microsoft Excel stores dates as sequential serial numbers so they can be used in calculations. By default, January 1, 1900 is serial number 1, and January 1, 2008 is serial number 39448 because it is 39,448 days after January 1, 1900. Microsoft Excel for the Macintosh uses a different date system as its default.
    
     **Note**  Visual Basic for Applications (VBA) calculates serial dates differently than Excel. In VBA, serial number 1 is December 31, 1899, rather than January 1, 1900. 
- All arguments are truncated to integers.
    
- If start_date or end_date are not valid dates, YEARFRAC returns the #VALUE! error value.
    
- If basis < 0 or if basis > 4, YEARFRAC returns the #NUM! error value.
    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

