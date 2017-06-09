---
title: DateValue Function
keywords: vblr6.chm1008889
f1_keywords:
- vblr6.chm1008889
ms.prod: office
ms.assetid: 8c9bd3d6-1614-eeb0-0714-4730eeeb1b95
ms.date: 06/08/2017
---


# DateValue Function



Returns a  **Variant** ( **Date** ).
 **Syntax**
 **DateValue(**_date_**)**
The required  _date_[argument](vbe-glossary.md) is normally a[string expression](vbe-glossary.md) representing a date from January 1, 100 through December 31, 9999. However, _date_ can also be any[expression](vbe-glossary.md) that can represent a date, a time, or both a date and time, in that range.
 **Remarks**
If  _date_ is a string that includes only numbers separated by valid[date separators](vbe-glossary.md),  **DateValue** recognizes the order for month, day, and year according to the Short Date format you specified for your system. **DateValue** also recognizes unambiguous dates that contain month names, either in long or abbreviated form. For example, in addition to recognizing 12/30/1991 and 12/30/91, **DateValue** also recognizes December 30, 1991 and Dec 30, 1991.
If the year part of  _date_ is omitted, **DateValue** uses the current year from your computer's system date.
If the  _date_ argument includes time information, **DateValue** doesn't return it. However, if _date_ includes invalid time information (such as "89:98"), an error occurs.

 **Note**  For  _date_, if the **Calendar** property setting is Gregorian, the supplied date must be Gregorian. If the calendar is Hijri, the supplied date must be Hijri. If the supplied date is Hijri, the argument _date_ is a **String** representing a date from 1/1/100 (Gregorian Aug 2, 718) through 4/3/9666 (Gregorian Dec 31, 9999).


## Example

This example uses the  **DateValue** function to convert a string to a date. You can also use date literals to directly assign a date to a **Variant** or **Date** variable, for example, MyDate = #2/12/69#.


```vb
Dim MyDate
MyDate = DateValue("February 12, 1969")    ' Return a date.


```


