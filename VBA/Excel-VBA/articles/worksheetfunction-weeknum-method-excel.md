---
title: WorksheetFunction.WeekNum Method (Excel)
keywords: vbaxl10.chm137341
f1_keywords:
- vbaxl10.chm137341
ms.prod: excel
api_name:
- Excel.WorksheetFunction.WeekNum
ms.assetid: 9a99ad5a-76ba-da98-34d9-b5ee09647b10
ms.date: 06/08/2017
---


# WorksheetFunction.WeekNum Method (Excel)

Returns a number that indicates where the week falls numerically within a year.


## Syntax

 _expression_ . **WeekNum**( **_Arg1_** , **_Arg2_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Serial_num - a date within the week. Dates should be entered by using the DATE function, or as results of other formulas or functions. For example, use DATE(2008,5,23) for the 23rd day of May, 2008. Problems can occur if dates are entered as text.|
| _Arg2_|Optional| **Variant**|Return_type - a number that determines on which day the week begins. The default is 1.|

### Return Value

Double


## Remarks


 **Important**   The WEEKNUM function considers the week containing January 1 to be the first week of the year. However, there is a European standard that defines the first week as the one with the majority of days (four or more) falling in the new year. This means that for years in which there are three days or less in the first week of January, the WEEKNUM function returns week numbers that are incorrect according to the European standard.



|**Return_type**|**Week Begins**|
|:-----|:-----|
|1|Week begins on Sunday. Weekdays are numbered 1 through 7.|
|2|Week begins on Monday. Weekdays are numbered 1 through 7.|
Microsoft Excel stores dates as sequential serial numbers so they can be used in calculations. By default, January 1, 1900 is serial number 1, and January 1, 2008 is serial number 39448 because it is 39,448 days after January 1, 1900. Microsoft Excel for the Macintosh uses a different date system as its default.


 **Note**  Visual Basic for Applications (VBA) calculates serial dates differently than Excel. In VBA, serial number 1 is December 31, 1899, rather than January 1, 1900. 


## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

