---
title: WeekdayName Function
keywords: vblr6.chm1008932
f1_keywords:
- vblr6.chm1008932
ms.prod: office
ms.assetid: 84a92bec-1e65-4f97-fdf9-cd524dd04081
ms.date: 06/08/2017
---


# WeekdayName Function



 **Description**
Returns a string indicating the specified day of the week.
 **Syntax**
 **WeekdayName(**_weekday_**,**_abbreviate_**,**_firstdayofweek_**)**
The  **WeekdayName** function syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _weekday_|Required. The numeric designation for the day of the week. Numeric value of each day depends on setting of the  _firstdayofweek_ setting.|
| _abbreviate_|Optional.  **Boolean** value that indicates if the weekday name is to be abbreviated. If omitted, the default is **False**, which means that the weekday name is not abbreviated.|
| _firstdayofweek_|Optional. Numeric value indicating the first day of the week. See Settings section for values.|
 **Settings**
The  _firstdayofweek_ argument can have the following values:


|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
|**vbUseSystem**|0|Use National Language Support (NLS) API setting.|
|**vbSunday**|1|Sunday (default)|
|**vbMonday**|2|Monday|
|**vbTuesday**|3|Tuesday|
|**vbWednesday**|4|Wednesday|
|**vbThursday**|5|Thursday|
|**vbFriday**|6|Friday|
|**vbSaturday**|7|Saturday|

