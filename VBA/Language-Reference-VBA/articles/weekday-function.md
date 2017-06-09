---
title: Weekday Function
keywords: vblr6.chm1009058
f1_keywords:
- vblr6.chm1009058
ms.prod: office
ms.assetid: 4e6197a7-5c55-e5cd-5164-ce1d31a9f80c
ms.date: 06/08/2017
---


# Weekday Function



Returns a  **Variant** ( **Integer** ) containing a whole number representing the day of the week.
 **Syntax**
 **Weekday(**_date_, [ **_firstdayofweek_** ] **)**
The  **Weekday** function syntax has these[named arguments](vbe-glossary.md):


|**Part**|**Description**|
|:-----|:-----|
|**_date_**|Required. [Variant](vbe-glossary.md), [numeric expression](vbe-glossary.md), [string expression](vbe-glossary.md), or any combination, that can represent a date. If  **_date_** contains[Null](vbe-glossary.md),  **Null** is returned.|
|**_firstdayofweek_**|Optional. A [constant](vbe-glossary.md) that specifies the first day of the week. If not specified, **vbSunday** is assumed.|
 **Settings**
The  **_firstdayofweek_** argument has these settings:


|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
|**vbUseSystem**|0|Use the NLS API setting.|
|**vbSunday**|1|Sunday (default)|
|**vbMonday**|2|Monday|
|**vbTuesday**|3|Tuesday|
|**vbWednesday**|4|Wednesday|
|**vbThursday**|5|Thursday|
|**vbFriday**|6|Friday|
|**vbSaturday**|7|Saturday|
 **Return Values**
The  **Weekday** function can return any of these values:


|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
|**vbSunday**|1|Sunday|
|**vbMonday**|2|Monday|
|**vbTuesday**|3|Tuesday|
|**vbWednesday**|4|Wednesday|
|**vbThursday**|5|Thursday|
|**vbFriday**|6|Friday|
|**vbSaturday**|7|Saturday|
 **Remarks**
If the  **Calendar** property setting is Gregorian, the returned integer represents the Gregorian day of the week for the date argument. If the calendar is Hijri, the returned integer represents the Hijri day of the week for the date argument. For Hijri dates, the argument number is any numeric expression that can represent a date and/or time from 1/1/100 (Gregorian Aug 2, 718) through 4/3/9666 (Gregorian Dec 31, 9999).

## Example

This example uses the  **Weekday** function to obtain the day of the week from a specified date.


```vb
Dim MyDate, MyWeekDay
MyDate = #February 12, 1969#    ' Assign a date.
MyWeekDay = Weekday(MyDate)    ' MyWeekDay contains 4 because 
    ' MyDate represents a Wednesday.


```


