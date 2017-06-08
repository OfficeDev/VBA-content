---
title: Month Function
keywords: vblr6.chm1008977
f1_keywords:
- vblr6.chm1008977
ms.prod: office
ms.assetid: d0b3cfef-d192-166d-2dcf-c60b08213327
ms.date: 06/08/2017
---


# Month Function



Returns a  **Variant** ( **Integer** ) specifying a whole number between 1 and 12, inclusive, representing the month of the year.
 **Syntax**
 **Month(**_date_**)**
The required  _date_[argument](vbe-glossary.md) is any[Variant](vbe-glossary.md), [numeric expression](vbe-glossary.md), [string expression](vbe-glossary.md), or any combination, that can represent a date. If  _date_ contains[Null](vbe-glossary.md),  **Null** is returned.

 **Note**  If the  **Calendar** property setting is Gregorian, the returned integer represents the Gregorian day of the week for the date argument. If the calendar is Hijri, the returned integer represents the Hijri day of the week for the date argument. For Hijri dates, the argument number is any numeric expression that can represent a date and/or time from 1/1/100 (Gregorian Aug 2, 718) through 4/3/9666 (Gregorian Dec 31, 9999).


## Example

This example uses the  **Month** function to obtain the month from a specified date. In the development environment, the date literal is displayed in short date format using the locale settings of your code.


```vb
Dim MyDate, MyMonth
MyDate = #February 12, 1969#    ' Assign a date.
MyMonth = Month(MyDate)    ' MyMonth contains 2.
```


