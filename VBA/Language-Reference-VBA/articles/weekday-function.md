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


| <strong>Part</strong>                    | <strong>Description</strong>                                                                                                                                                                                                                                              |
|:-----------------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| <strong><em>date</em></strong>           | Required. [Variant](vbe-glossary.md), [numeric expression](vbe-glossary.md), [string expression](vbe-glossary.md), or any combination, that can represent a date. If  <strong><em>date</em></strong> contains[Null](vbe-glossary.md),  <strong>Null</strong> is returned. |
| <strong><em>firstdayofweek</em></strong> | Optional. A [constant](vbe-glossary.md) that specifies the first day of the week. If not specified, <strong>vbSunday</strong> is assumed.                                                                                                                                 |

 **Settings**
The  **_firstdayofweek_** argument has these settings:


| <strong>Constant</strong>    | <strong>Value</strong> | <strong>Description</strong> |
|:-----------------------------|:-----------------------|:-----------------------------|
| <strong>vbUseSystem</strong> | 0                      | Use the NLS API setting.     |
| <strong>vbSunday</strong>    | 1                      | Sunday (default)             |
| <strong>vbMonday</strong>    | 2                      | Monday                       |
| <strong>vbTuesday</strong>   | 3                      | Tuesday                      |
| <strong>vbWednesday</strong> | 4                      | Wednesday                    |
| <strong>vbThursday</strong>  | 5                      | Thursday                     |
| <strong>vbFriday</strong>    | 6                      | Friday                       |
| <strong>vbSaturday</strong>  | 7                      | Saturday                     |

 **Return Values**
The  **Weekday** function can return any of these values:


| <strong>Constant</strong>    | <strong>Value</strong> | <strong>Description</strong> |
|:-----------------------------|:-----------------------|:-----------------------------|
| <strong>vbSunday</strong>    | 1                      | Sunday                       |
| <strong>vbMonday</strong>    | 2                      | Monday                       |
| <strong>vbTuesday</strong>   | 3                      | Tuesday                      |
| <strong>vbWednesday</strong> | 4                      | Wednesday                    |
| <strong>vbThursday</strong>  | 5                      | Thursday                     |
| <strong>vbFriday</strong>    | 6                      | Friday                       |
| <strong>vbSaturday</strong>  | 7                      | Saturday                     |

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


