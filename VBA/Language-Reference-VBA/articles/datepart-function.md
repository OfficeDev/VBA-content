---
title: DatePart Function
keywords: vblr6.chm1012951
f1_keywords:
- vblr6.chm1012951
ms.prod: office
ms.assetid: 65476ecc-c1d6-333e-b8b5-417a96373594
ms.date: 06/08/2017
---


# DatePart Function



Returns a  **Variant** ( **Integer** ) containing the specified part of a given date.
 **Syntax**
 **DatePart( _interval,_** **_date_** [ **_,firstdayofweek_** [ **_,_** **_firstweekofyear_** ]] **)**
The  **DatePart** function syntax has these[named arguments](vbe-glossary.md):


| <strong>Part</strong>                     | <strong>Description</strong>                                                                                                                          |
|:------------------------------------------|:------------------------------------------------------------------------------------------------------------------------------------------------------|
| <strong><em>interval</em></strong>        | Required. [String expression](vbe-glossary.md) that is the interval of time you want to return.                                                       |
| <strong><em>date</em></strong>            | Required.  <strong>Variant</strong> ( <strong>Date</strong> ) value that you want to evaluate.                                                        |
| <strong><em>firstdayofweek</em></strong>  | Optional. A [constant](vbe-glossary.md) that specifies the first day of the week. If not specified, Sunday is assumed.                                |
| <strong><em>firstweekofyear</em></strong> | Optional. A constant that specifies the first week of the year. If not specified, the first week is assumed to be the week in which January 1 occurs. |

 **Settings**
The  **_interval_**[argument](vbe-glossary.md) has these settings:


| <strong>Setting</strong> | <strong>Description</strong> |
|:-------------------------|:-----------------------------|
| yyyy                     | Year                         |
| q                        | Quarter                      |
| m                        | Month                        |
| y                        | Day of year                  |
| d                        | Day                          |
| w                        | Weekday                      |
| ww                       | Week                         |
| h                        | Hour                         |
| n                        | Minute                       |
| s                        | Second                       |

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

The  **_firstweekofyear_** argument has these settings:


| <strong>Constant</strong>        | <strong>Value</strong> | <strong>Description</strong>                                           |
|:---------------------------------|:-----------------------|:-----------------------------------------------------------------------|
| <strong>vbUseSystem</strong>     | 0                      | Use the NLS API setting.                                               |
| <strong>vbFirstJan1</strong>     | 1                      | Start with week in which January 1 occurs (default).                   |
| <strong>vbFirstFourDays</strong> | 2                      | Start with the first week that has at least four days in the new year. |
| <strong>vbFirstFullWeek</strong> | 3                      | Start with first full week of the year.                                |

 **Remarks**
You can use the  **DatePart** function to evaluate a date and return a specific interval of time. For example, you might use **DatePart** to calculate the day of the week or the current hour.
The  **_firstdayofweek_** argument affects calculations that use the "w" and "ww" interval symbols.
If  _date_ is a[date literal](vbe-glossary.md), the specified year becomes a permanent part of that date. However, if  _date_ is enclosed in double quotation marks (" "), and you omit the year, the current year is inserted in your code each time the _date_ expression is evaluated. This makes it possible to write code that can be used in different years.

 **Note**  For  _date_, if the **Calendar** property setting is Gregorian, the supplied date must be Gregorian. If the calendar is Hijri, the supplied date must be Hijri.

The returned date part is in the time period units of the current Arabic calendar. For example, if the current calendar is Hijri and the date part to be returned is the year, the year value is a Hijri year.

## Example

This example takes a date and, using the  **DatePart** function, displays the quarter of the year in which it occurs.


```vb
Dim TheDate As Date    ' Declare variables.
Dim Msg    
TheDate = InputBox("Enter a date:")
Msg = "Quarter: " &; DatePart("q", TheDate)
MsgBox Msg
```


