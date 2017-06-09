---
title: WorksheetFunction.Days360 Method (Excel)
keywords: vbaxl10.chm137160
f1_keywords:
- vbaxl10.chm137160
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Days360
ms.assetid: cc8b607d-348f-0fa7-70e4-3ddb9b83f6b8
ms.date: 06/08/2017
---


# WorksheetFunction.Days360 Method (Excel)

Returns the number of days between two dates based on a 360-day year (twelve 30-day months), which is used in some accounting calculations.


## Syntax

 _expression_ . **Days360**( **_Arg1_** , **_Arg2_** , **_Arg3_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1 - Arg2_|Required| **Variant**|The two dates between which you want to know the number of days. If start_date (Arg1) occurs after end_date (Arg2), Days360 returns a negative number. Dates should be entered by using the DATE function, or as results of other formulas or functions.|
| _Arg3_|Optional| **Variant**|A boolean value that specifies whether to use the U.S. or European method in the calculation.|

### Return Value

Double


## Remarks

 Use this function to help compute payments if your accounting system is based on twelve 30-day months.

The following tables contains the values for  _Arg3_ .



|**Method**|**Defined**|
|:-----|:-----|
|FALSE or omitted|U.S. (NASD) method. If the starting date is the 31st of a month, it becomes equal to the 30th of the same month. If the ending date is the 31st of a month and the starting date is earlier than the 30th of a month, the ending date becomes equal to the 1st of the next month; otherwise the ending date becomes equal to the 30th of the same month.|
|TRUE|European method. Starting dates and ending dates that occur on the 31st of a month become equal to the 30th of the same month.|

 **Caution**  When you use the DAYS360 function to calculate the number of days between two dates, an unexpected value is returned. For example, when you use the DAYS360 function with a start date of February 28 and with an end date of March 28, a value of 28 days is returned. You expect a value of 30 days to be returned for every full month.To work around this behavior, use the following formula: =DAYS360(start_date,end_date,IF(method=TRUE,TRUE,IF(AND(method=FALSE,MONTH(start_date)=2,DAY(start_date)>=28,MONTH(end_date)=2,DAY(end_date)>=28),TRUE,FALSE)))

Microsoft Excel stores dates as sequential serial numbers so they can be used in calculations. By default, January 1, 1900 is serial number 1, and January 1, 2008 is serial number 39448 because it is 39,448 days after January 1, 1900. Microsoft Excel for the Macintosh uses a different date system as its default. 


 **Note**  Visual Basic for Applications (VBA) calculates serial dates differently than Excel. In VBA, serial number 1 is December 31, 1899, rather than January 1, 1900. 


## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

