---
title: Day Function
keywords: vblr6.chm1008890
f1_keywords:
- vblr6.chm1008890
ms.prod: office
ms.assetid: 8d4d0b63-28d9-c6a3-bd96-3688b0f93a12
ms.date: 06/08/2017
---


# Day Function



Returns a  **Variant** ( **Integer** ) specifying a whole number between 1 and 31, inclusive, representing the day of the month.
 **Syntax**
 **Day(**_date_**)**
The required  _date_[argument](vbe-glossary.md) is any[Variant](vbe-glossary.md), [numeric expression](vbe-glossary.md), [string expression](vbe-glossary.md), or any combination, that can represent a date. If  _date_ contains[Null](vbe-glossary.md),  **Null** is returned.

 **Note**  If the  **Calendar** property setting is Gregorian, the returned integer represents the Gregorian day of the month for the date argument. If the calendar is Hijri, the returned integer represents the Hijri day of the month for the date argument.


## Example

This example uses the  **Day** function to obtain the day of the month from a specified date. In the development environment, the date literal is displayed in short format using the locale settings of your code.


```vb
Dim MyDate, MyDay
MyDate = #February 12, 1969#    ' Assign a date.
MyDay = Day(MyDate)    ' MyDay contains 12.


```


