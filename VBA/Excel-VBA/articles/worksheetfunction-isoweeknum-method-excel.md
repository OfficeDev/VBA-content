---
title: WorksheetFunction.IsoWeekNum Method (Excel)
keywords: vbaxl10.chm137457
f1_keywords:
- vbaxl10.chm137457
ms.prod: excel
ms.assetid: 8b643312-d9b9-c509-ca9f-c3d960ba012c
ms.date: 06/08/2017
---


# WorksheetFunction.IsoWeekNum Method (Excel)

Returns the ISO week number of the year for a given date. .


## Syntax

 _expression_ . **IsoWeekNum**_(Arg1,_ _Arg2)_

 _expression_ A variable that represents a[WorksheetFunction Object (Excel)](worksheetfunction-object-excel.md) object.


### Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required|DOUBLE|Date-time code used by Microsoft Excel for date and time calculation.|
| _Arg2_|Optional|VARIANT|This argument is not available in the function.|

### Return value

 **DOUBLE**


## Remarks

Returns the ordinal number of the [ISO8601] calendar week in the year for the given date. [ISO 8601](http://en.wikipedia.org/wiki/ISO_8601) defines the calendar week as a time interval of seven calendar days starting with a Monday, and the first calendar week of a year as the one that includes the first Thursday of that year.


## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

