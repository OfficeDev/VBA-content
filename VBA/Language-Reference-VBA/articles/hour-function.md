---
title: Hour Function
keywords: vblr6.chm1008939
f1_keywords:
- vblr6.chm1008939
ms.prod: office
ms.assetid: cf0800d1-6e26-71ad-ec8d-09e4876bf469
ms.date: 06/08/2017
---


# Hour Function



Returns a  **Variant** ( **Integer** ) specifying a whole number between 0 and 23, inclusive, representing the hour of the day.
 **Syntax**
 **Hour(**_time_**)**
The required  _time_[argument](vbe-glossary.md) is any[Variant](vbe-glossary.md), [numeric expression](vbe-glossary.md), [string expression](vbe-glossary.md), or any combination, that can represent a time. If  _time_ contains[Null](vbe-glossary.md),  **Null** is returned.

## Example

This example uses the  **Hour** function to obtain the hour from a specified time. In the development environment, the time literal is displayed in short time format using the locale settings of your code.


```vb
Dim MyTime, MyHour
MyTime = #4:35:17 PM#    ' Assign a time.
MyHour = Hour(MyTime)    ' MyHour contains 16.


```


