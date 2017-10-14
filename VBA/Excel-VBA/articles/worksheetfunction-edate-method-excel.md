---
title: WorksheetFunction.EDate Method (Excel)
keywords: vbaxl10.chm137325
f1_keywords:
- vbaxl10.chm137325
ms.prod: excel
api_name:
- Excel.WorksheetFunction.EDate
ms.assetid: c3f068c2-f6ef-bcb7-79db-e1de4348038c
ms.date: 06/08/2017
---


# WorksheetFunction.EDate Method (Excel)

Returns the serial number that represents the date that is the indicated number of months before or after a specified date (the start_date). Use EDATE to calculate maturity dates or due dates that fall on the same day of the month as the date of issue.


## Syntax

 _expression_ . **EDate**( **_Arg1_** , **_Arg2_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Start_date - a date that represents the start date. Dates should be entered by using the DATE function, or as results of other formulas or functions. For example, use DATE(2008,5,23) for the 23rd day of May, 2008. Problems can occur if dates are entered as text.|
| _Arg2_|Required| **Variant**|Months - the number of months before or after start_date. A positive value for months yields a future date; a negative value yields a past date.|

### Return Value

Double


## Remarks




- Microsoft Excel stores dates as sequential serial numbers so they can be used in calculations. By default, January 1, 1900 is serial number 1, and January 1, 2008 is serial number 39448 because it is 39,448 days after January 1, 1900. Microsoft Excel for the Macintosh uses a different date system as its default.
    
     **Note**  Visual Basic for Applications (VBA) calculates serial dates differently than Excel. In VBA, serial number 1 is December 31, 1899, rather than January 1, 1900. 
- If start_date is not a valid date, EDATE returns the #VALUE! error value.
    
- If months is not an integer, it is truncated.
    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

