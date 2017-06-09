---
title: Second Function
keywords: vblr6.chm1009011
f1_keywords:
- vblr6.chm1009011
ms.prod: office
ms.assetid: fef87486-ccda-23e7-04a5-5e484ce66543
ms.date: 06/08/2017
---


# Second Function



Returns a  **Variant** ( **Integer** ) specifying a whole number between 0 and 59, inclusive, representing the second of the minute.
 **Syntax**
 **Second(**_time_**)**
The required  _time_[argument](vbe-glossary.md) is any[Variant](vbe-glossary.md), [numeric expression](vbe-glossary.md), [string expression](vbe-glossary.md), or any combination, that can represent a time. If  _time_ contains[Null](vbe-glossary.md),  **Null** is returned.

## Example

This example uses the  **Second** function to obtain the second of the minute from a specified time. In the development environment, the time literal is displayed in short time format using the locale settings of your code.


```vb
Dim MyTime, MySecond
MyTime = #4:35:17 PM#    ' Assign a time.
MySecond = Second(MyTime)    ' MySecond contains 17.


```


