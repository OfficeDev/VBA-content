---
title: DateDiff Function
keywords: vblr6.chm1012950
f1_keywords:
- vblr6.chm1012950
ms.prod: office
ms.assetid: 15c9df5f-1403-b6a5-71b9-611e9820d804
ms.date: 06/08/2017
---


# DateDiff Function



Returns a  **Variant** ( **Long** ) specifying the number of time intervals between two specified dates.
 **Syntax**
 **DateDiff( _interval, date1, date2_** [ **_, firstdayofweek_** [ **,** **_firstweekofyear_** ]] **)**
The  **DateDiff** function syntax has these[named arguments](vbe-glossary.md):


|**Part**|**Description**|
|:-----|:-----|
|**_interval_**|Required. [String expression](vbe-glossary.md) that is the interval of time you use to calculate the difference between **_date1_** and **_date2_**.|
|**_date1_**, **_date2_**|Required;  **Variant** ( **Date** ). Two dates you want to use in the calculation.|
|**_firstdayofweek_**|Optional. A [constant](vbe-glossary.md) that specifies the first day of the week. If not specified, Sunday is assumed.|
|**_firstweekofyear_**|Optional. A constant that specifies the first week of the year. If not specified, the first week is assumed to be the week in which January 1 occurs.|
 **Settings**
The  **_interval_**[argument](vbe-glossary.md) has these settings:


|**Setting**|**Description**|
|:-----|:-----|
|yyyy|Year|
|q|Quarter|
|m|Month|
|y|Day of year|
|d|Day|
|w|Weekday|
|ww|Week|
|h|Hour|
|n|Minute|
|s|Second|
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


|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
|**vbUseSystem**|0|Use the NLS API setting.|
|**vbFirstJan1**|1|Start with week in which January 1 occurs (default).|
|**vbFirstFourDays**|2|Start with the first week that has at least four days in the new year.|
|**vbFirstFullWeek**|3|Start with first full week of the year.|
 **Remarks**
You can use the  **DateDiff** function to determine how many specified time intervals exist between two dates. For example, you might use **DateDiff** to calculate the number of days between two dates, or the number of weeks between today and the end of the year.
To calculate the number of days between  **_date1_** and **_date2_**, you can use either Day of year ("y") or Day ("d"). When **_interval_** is Weekday ("w"), **DateDiff** returns the number of weeks between the two dates. If **_date1_** falls on a Monday, **DateDiff** counts the number of Mondays until **_date2_**. It counts **_date2_** but not **_date1_**. If **_interval_** is Week ("ww"), however, the **DateDiff** function returns the number of calendar weeks between the two dates. It counts the number of Sundays between **_date1_** and **_date2_**. **DateDiff** counts **_date2_** if it falls on a Sunday; but it doesn't count **_date1_**, even if it does fall on a Sunday.
If  **_date1_** refers to a later point in time than **_date2_**, the **DateDiff** function returns a negative number.
The  **_firstdayofweek_** argument affects calculations that use the "w" and "ww" interval symbols.
If  **_date1_** or **_date2_** is a[date literal](vbe-glossary.md), the specified year becomes a permanent part of that date. However, if  **_date1_** or _date2_ is enclosed in double quotation marks (" "), and you omit the year, the current year is inserted in your code each time the **_date1_** or _date2_ expression is evaluated. This makes it possible to write code that can be used in different years.
When comparing December 31 to January 1 of the immediately succeeding year,  **DateDiff** for Year ("yyyy") returns 1 even though only a day has elapsed.

 **Note**  For  **_date1_** and **_date2_**, if the **Calendar** property setting is Gregorian, the supplied date must be Gregorian. If the calendar is Hijri, the supplied date must be Hijri.


## Example

This example uses the  **DateDiff** function to display the number of days between a given date and today.


```vb
Dim TheDate As Date    ' Declare variables.
Dim Msg
TheDate = InputBox("Enter a date")
Msg = "Days from today: " &; DateDiff("d", Now, TheDate)
MsgBox Msg


```


