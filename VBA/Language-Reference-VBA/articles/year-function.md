---
title: Year Function
keywords: vblr6.chm1009063
f1_keywords:
- vblr6.chm1009063
ms.prod: office
ms.assetid: c82b30dd-a8ce-c213-3619-7de33278a3c8
ms.date: 06/08/2017
---


# Year Function



Returns a  **Variant** ( **Integer** ) containing a whole number representing the year.
 **Syntax**
 **Year(**_date_**)**
The required  _date_[argument](vbe-glossary.md) is any[Variant](vbe-glossary.md), [numeric expression](vbe-glossary.md), [string expression](vbe-glossary.md), or any combination, that can represent a date. If  _date_ contains[Null](vbe-glossary.md),  **Null** is returned.

 **Note**  If the  **Calendar** property setting is Gregorian, the returned integer represents the Gregorian year for the date argument. If the calendar is Hijri, the returned integer represents the Hijri year for the date argument. For Hijri dates, the argument number is any numeric expression that can represent a date and/or time from 1/1/100 (Gregorian Aug 2, 718) through 4/3/9666 (Gregorian Dec 31, 9999).


## Example

This example uses the  **Year** function to obtain the year from a specified date. In the development environment, the date literal is displayed in short date format using the locale settings of your code.


```vb
Dim MyDate, MyYear
MyDate = #February 12, 1969#    ' Assign a date.
MyYear = Year(MyDate)    ' MyYear contains 1969.


```


