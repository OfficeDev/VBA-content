---
title: DateSerial Function
keywords: vblr6.chm1008888
f1_keywords:
- vblr6.chm1008888
ms.prod: office
ms.assetid: 3aae4616-ab44-5e57-ba10-1d5ca1659c6e
ms.date: 06/08/2017
---


# DateSerial Function



Returns a  **Variant** ( **Date** ) for a specified year, month, and day.
 **Syntax**
 **DateSerial( _year_, _month_, _day_ )**
The  **DateSerial** function syntax has these[named arguments](vbe-glossary.md):


|**Part**|**Description**|
|:-----|:-----|
|**_year_**|Required;  **Integer**. Number between 100 and 9999, inclusive, or a[numeric expression](vbe-glossary.md).|
|**_month_**|Required;  **Integer**. Any numeric expression.|
|**_day_**|Required;  **Integer**. Any numeric expression.|
 **Remarks**
To specify a date, such as December 31, 1991, the range of numbers for each  **DateSerial**[argument](vbe-glossary.md) should be in the accepted range for the unit; that is, 1-31 for days and 1-12 for months. However, you can also specify relative dates for each argument using any numeric expression that represents some number of days, months, or years before or after a certain date.
The following example uses numeric expressions instead of absolute date numbers. Here the  **DateSerial** function returns a date that is the day before the first day ( `1 - 1`), two months before August ( `8 - 2`), 10 years before 1990 (), two months before August ( `8 - 2`), 10 years before 1990 ( `1990 - 10`); in other words, May 31, 1980.
Under Windows 98 or Windows 2000, two digit years for the  **_year_** argument are interpreted based on user-defined machine settings. The default settings are that values between 0 and 29, inclusive, are interpreted as the years 2000-2029. The default values between 30 and 99 are interpreted as the years 1930-1999. For all other **_year_** arguments, use a four-digit year (for example, 1800).
Earlier versions of Windows interpret two-digit years based on the defaults described above. To be sure the function returns the proper value, use a four-digit year.
When any argument exceeds the accepted range for that argument, it increments to the next larger unit as appropriate. For example, if you specify 35 days, it is evaluated as one month and some number of days, depending on where in the year it is applied. If any single argument is outside the range -32,768 to 32,767, an error occurs. If the date specified by the three arguments falls outside the acceptable range of dates, an error occurs.

 **Note**  For  _year, month,_ and _day_, if the **Calendar** property setting is Gregorian, the supplied value is assumed to be Gregorian. If the **Calendar** property setting is Hijri, the supplied value is assumed to be Hijri.

The returned date part is in the time period units of the current Visual Basic calendar. For example, if the current calendar is Hijri and the date part to be returned is the year, the year value is a Hijri year. For the argument  **_year_**, values between 0 and 99, inclusive, are interpreted as the years 1400-1499. For all other **_year_** values, use the complete four-digit year (for example, 1520).

## Example

This example uses the  **DateSerial** function to return the date for the specified year, month, and day.


```vb
Dim MyDate
' MyDate contains the date for February 12, 1969.
MyDate = DateSerial(1969, 2, 12)    ' Return a date.


```


